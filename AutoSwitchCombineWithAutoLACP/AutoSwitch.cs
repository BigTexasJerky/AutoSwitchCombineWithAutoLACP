using HarmonyLib;
using Il2CppInterop.Runtime.Injection;
using MelonLoader;
using MelonLoader.Utils;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using UnityEngine;
using Component = UnityEngine.Component;

[assembly: MelonInfo(typeof(AutoSwitch.AutoSwitchMod), "AutoSwitch", "2.10.0", "Big Texas Jerky")]
[assembly: MelonGame("Waseku", "Data Center")]

namespace AutoSwitch
{
    public sealed class AutoSwitchMod : MelonMod
    {
        private const float ScanIntervalSeconds = 8.0f;

        private const float SameFabricXZTolerance = 0.40f;
        private const float AdjacentYTolerance = 0.040f;
        private const float FallbackExpectedStep = 0.0444f;

        private const float OriginXZRejectRadius = 1.25f;
        private const float MinUsefulWorldY = 1.0f;
        private const float MaxUsefulWorldY = 4.5f;

        private static readonly Regex TrailingRuntimeIdRegex =
            new Regex(@"_[\-]?\d+$", RegexOptions.Compiled);

        private static string DebugFolderPath =>
            Path.Combine(MelonEnvironment.ModsDirectory, "AutoSwitch");

        private static string DebugLogPath =>
            Path.Combine(DebugFolderPath, "autoswitch-debug.log");

        private static readonly HashSet<string> AllowedSwitchModels = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Switch32xQSFP",
            "Switch4xQSXP16xSFP",
            "Switch4xSFP",
            "Switch16CU"
        };

        private static readonly Dictionary<string, RegisteredSwitchInfo> RegisteredSwitches =
            new Dictionary<string, RegisteredSwitchInfo>(StringComparer.OrdinalIgnoreCase);

        private static readonly Dictionary<int, RegisteredCableInfo> RegisteredCables =
            new Dictionary<int, RegisteredCableInfo>();

        private static readonly Dictionary<string, float> FabricIdToSpeed =
            new Dictionary<string, float>(StringComparer.OrdinalIgnoreCase);

        private static readonly Dictionary<string, HashSet<int>> FabricIdToBundleCableIds =
            new Dictionary<string, HashSet<int>>(StringComparer.OrdinalIgnoreCase);

        private static readonly HashSet<int> WatchedCableIds =
            new HashSet<int>();

        private static readonly HashSet<string> WatchedDevices =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private static readonly HashSet<string> LoggedPatches =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private static readonly HashSet<string> LoggedBundles =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private static readonly HashSet<string> LoggedProbeEvents =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private static Type _networkMapType;

        private float _nextScanTime;
        private string _lastSummary = string.Empty;

        public override void OnInitializeMelon()
        {
            ClassInjector.RegisterTypeInIl2Cpp<FabricGroupTag>();

            Directory.CreateDirectory(DebugFolderPath);
            File.WriteAllText(DebugLogPath, "[" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "] AutoSwitch 2.10 debug log started." + Environment.NewLine);

            InstallNativePatches();

            MelonLogger.Msg("[AutoSwitch] v2.10.0 active. Route/path probe build.");
        }

        public override void OnSceneWasLoaded(int buildIndex, string sceneName)
        {
            if (buildIndex == 0)
                return;

            _nextScanTime = 0f;
            _lastSummary = string.Empty;

            RegisteredSwitches.Clear();
            RegisteredCables.Clear();
            FabricIdToSpeed.Clear();
            FabricIdToBundleCableIds.Clear();
            WatchedCableIds.Clear();
            WatchedDevices.Clear();
            LoggedBundles.Clear();
            LoggedProbeEvents.Clear();

            LogToFile("Scene loaded: " + sceneName + " (" + buildIndex.ToString(CultureInfo.InvariantCulture) + ")");
        }

        public override void OnUpdate()
        {
            try
            {
                if (Time.time < _nextScanTime)
                    return;

                _nextScanTime = Time.time + ScanIntervalSeconds;
                RunRoutePathProbePass();
            }
            catch (Exception ex)
            {
                MelonLogger.Error("[AutoSwitch] OnUpdate exception: " + ex);
                LogToFile("OnUpdate exception: " + ex);
                _nextScanTime = Time.time + 10.0f;
            }
        }

        private void RunRoutePathProbePass()
        {
            RefreshRegisteredSwitchPositions();

            List<RegisteredSwitchInfo> liveSwitches = RegisteredSwitches.Values
                .Where(s => s != null && s.HasWorldPosition && IsAllowedSwitchModel(s.ModelName))
                .Where(IsLikelyPlacedSwitch)
                .ToList();

            List<List<RegisteredSwitchInfo>> fabrics = BuildRegistryFabrics(liveSwitches)
                .Where(f => f.Count >= 2)
                .OrderByDescending(f => f.Count)
                .ThenByDescending(f => f.Sum(x => x.EstimatedCapacityGbps))
                .ToList();

            DetectRemoteDeviceBundles(fabrics);

            int adjacencyPairs = fabrics.Sum(f => Math.Max(0, f.Count - 1));

            string summary =
                "SCAN SUMMARY | registeredSwitches=" + RegisteredSwitches.Count.ToString(CultureInfo.InvariantCulture) +
                " | registeredCables=" + RegisteredCables.Count.ToString(CultureInfo.InvariantCulture) +
                " | liveSwitches=" + liveSwitches.Count.ToString(CultureInfo.InvariantCulture) +
                " | activeFabrics=" + fabrics.Count.ToString(CultureInfo.InvariantCulture) +
                " | adjacentPairs=" + adjacencyPairs.ToString(CultureInfo.InvariantCulture) +
                " | watchedCableIds=" + WatchedCableIds.Count.ToString(CultureInfo.InvariantCulture) +
                " | watchedDevices=" + WatchedDevices.Count.ToString(CultureInfo.InvariantCulture);

            if (!string.Equals(summary, _lastSummary, StringComparison.Ordinal))
            {
                _lastSummary = summary;
                MelonLogger.Msg("[AutoSwitch] " + summary);
                LogToFile(summary);

                int fabricIndex = 1;
                foreach (List<RegisteredSwitchInfo> fabric in fabrics)
                {
                    string fabricId = "FABRIC-" + fabricIndex.ToString("000", CultureInfo.InvariantCulture);
                    float speed = FabricIdToSpeed.ContainsKey(fabricId) ? FabricIdToSpeed[fabricId] : fabric.Sum(x => x.EstimatedCapacityGbps);
                    int tracked = FabricIdToBundleCableIds.ContainsKey(fabricId) ? FabricIdToBundleCableIds[fabricId].Count : 0;

                    string members = string.Join(", ",
                        fabric.Select(x => x.DeviceName + "@(" +
                                           x.WorldPosition.x.ToString("0.###", CultureInfo.InvariantCulture) + "," +
                                           x.WorldPosition.y.ToString("0.###", CultureInfo.InvariantCulture) + "," +
                                           x.WorldPosition.z.ToString("0.###", CultureInfo.InvariantCulture) + ")[" +
                                           x.ModelName + "]"));

                    LogToFile("FABRIC | id=" + fabricId +
                              " | members=" + fabric.Count.ToString(CultureInfo.InvariantCulture) +
                              " | estCapacityGbps=" + speed.ToString("0.##", CultureInfo.InvariantCulture) +
                              " | bundleCableIds=" + tracked.ToString(CultureInfo.InvariantCulture) +
                              " | switches=" + members);

                    if (tracked > 0)
                    {
                        string cableList = string.Join(",", FabricIdToBundleCableIds[fabricId].OrderBy(x => x));
                        LogToFile("FABRIC BUNDLE CABLE MAP | fabricId=" + fabricId + " | cableIds=[" + cableList + "]");
                    }

                    fabricIndex++;
                }
            }
        }

        private void InstallNativePatches()
        {
            try
            {
                _networkMapType = AccessTools.TypeByName("Il2Cpp.NetworkMap");

                if (_networkMapType != null)
                {
                    PatchIfFound(_networkMapType, "RegisterSwitch", nameof(NetworkMap_RegisterSwitch_Postfix), patchPostfix: true);
                    PatchIfFound(_networkMapType, "RegisterCableConnection", nameof(NetworkMap_RegisterCableConnection_Postfix), patchPostfix: true);

                    PatchIfFound(_networkMapType, "FindAllRoutes",
                        nameof(NetworkMap_FindAllRoutes_Postfix),
                        patchPostfix: true,
                        patchPrefixName: nameof(NetworkMap_FindAllRoutes_Prefix));

                    PatchIfFound(_networkMapType, "FindPhysicalPath",
                        nameof(NetworkMap_FindPhysicalPath_Postfix),
                        patchPostfix: true,
                        patchPrefixName: nameof(NetworkMap_FindPhysicalPath_Prefix));

                    PatchIfFound(_networkMapType, "UpdateCustomerServerCountAndSpeed",
                        nameof(NetworkMap_UpdateCustomerServerCountAndSpeed_Postfix),
                        patchPostfix: true,
                        patchPrefixName: nameof(NetworkMap_UpdateCustomerServerCountAndSpeed_Prefix));

                    PatchIfFound(_networkMapType, "GetLACPGroupForCable",
                        nameof(NetworkMap_GetLACPGroupForCable_Postfix),
                        patchPostfix: true);

                    PatchIfFound(_networkMapType, "GetLACPGroupBetween",
                        nameof(NetworkMap_GetLACPGroupBetween_Postfix),
                        patchPostfix: true);
                }
            }
            catch (Exception ex)
            {
                LogToFile("InstallNativePatches failed: " + ex);
            }
        }

        private void PatchIfFound(Type type, string methodName, string postfixName = null, bool patchPostfix = false, string patchPrefixName = null)
        {
            try
            {
                MethodInfo target = AccessTools.Method(type, methodName);
                if (target == null)
                {
                    LogToFile("PATCH MISS | " + type.FullName + "." + methodName);
                    return;
                }

                HarmonyMethod prefix = null;
                HarmonyMethod postfix = null;

                if (!string.IsNullOrWhiteSpace(patchPrefixName))
                {
                    MethodInfo prefixMethod = typeof(AutoSwitchMod).GetMethod(patchPrefixName, BindingFlags.NonPublic | BindingFlags.Static);
                    if (prefixMethod != null)
                        prefix = new HarmonyMethod(prefixMethod);
                }

                if (patchPostfix && !string.IsNullOrWhiteSpace(postfixName))
                {
                    MethodInfo postfixMethod = typeof(AutoSwitchMod).GetMethod(postfixName, BindingFlags.NonPublic | BindingFlags.Static);
                    if (postfixMethod != null)
                        postfix = new HarmonyMethod(postfixMethod);
                }

                HarmonyInstance.Patch(target, prefix: prefix, postfix: postfix);
                LogPatch("Patched " + type.FullName + "." + methodName);
            }
            catch (Exception ex)
            {
                LogToFile("PATCH FAIL | " + type.FullName + "." + methodName + " | " + ex);
            }
        }

        private static void LogPatch(string message)
        {
            if (LoggedPatches.Add(message))
                LogToFile(message);
        }

        private static void NetworkMap_RegisterSwitch_Postfix(object __0)
        {
            try
            {
                if (__0 == null)
                    return;

                RegisteredSwitchInfo info = BuildRegisteredSwitchInfo(__0);
                if (info == null || string.IsNullOrWhiteSpace(info.DeviceName))
                    return;

                RegisteredSwitches[info.DeviceName] = info;

                LogToFile("REGISTER SWITCH | deviceName=" + info.DeviceName +
                          " | model=" + info.ModelName +
                          " | pos=(" +
                          info.WorldPosition.x.ToString("0.###", CultureInfo.InvariantCulture) + "," +
                          info.WorldPosition.y.ToString("0.###", CultureInfo.InvariantCulture) + "," +
                          info.WorldPosition.z.ToString("0.###", CultureInfo.InvariantCulture) + ")");
            }
            catch (Exception ex)
            {
                LogToFile("NetworkMap_RegisterSwitch_Postfix failed: " + ex);
            }
        }

        private static void NetworkMap_RegisterCableConnection_Postfix(
            int __0,
            Vector3 __1,
            Vector3 __2,
            object __3,
            object __4,
            string __5,
            string __6,
            int __7,
            int __8,
            string __9,
            string __10)
        {
            try
            {
                RegisteredCableInfo info = new RegisteredCableInfo();
                info.CableId = __0;
                info.DeviceA = (__5 ?? string.Empty).Trim();
                info.DeviceB = (__6 ?? string.Empty).Trim();
                info.ExtraA = (__9 ?? string.Empty).Trim();
                info.ExtraB = (__10 ?? string.Empty).Trim();
                info.PortA = __7;
                info.PortB = __8;
                info.EndpointObjectA = __3;
                info.EndpointObjectB = __4;
                info.ServerRootKeyA = ResolveServerRootKey(__3, __9, __5);
                info.ServerRootKeyB = ResolveServerRootKey(__4, __10, __6);

                RegisteredCables[__0] = info;

                LogToFile("REGISTER CABLE | cableId=" + info.CableId.ToString(CultureInfo.InvariantCulture) +
                          " | deviceA=" + info.DeviceA +
                          " | deviceB=" + info.DeviceB +
                          " | extraA=" + info.ExtraA +
                          " | extraB=" + info.ExtraB +
                          " | serverRootA=" + info.ServerRootKeyA +
                          " | serverRootB=" + info.ServerRootKeyB +
                          " | portA=" + info.PortA.ToString(CultureInfo.InvariantCulture) +
                          " | portB=" + info.PortB.ToString(CultureInfo.InvariantCulture));
            }
            catch (Exception ex)
            {
                LogToFile("NetworkMap_RegisterCableConnection_Postfix failed: " + ex);
            }
        }

        private static void NetworkMap_FindAllRoutes_Prefix(string __0, string __1)
        {
            try
            {
                if (!IsInterestingRouteRequest(__0, __1))
                    return;

                LogOnce("FindAllRoutesPRE|" + __0 + "|" + __1,
                    "ROUTE PATH | FindAllRoutes PRE | base=" + __0 + " | server=" + __1);
            }
            catch
            {
            }
        }

        private static void NetworkMap_FindAllRoutes_Postfix(string __0, string __1, object __result)
        {
            try
            {
                if (!IsInterestingRouteRequest(__0, __1))
                    return;

                LogToFile("ROUTE PATH | FindAllRoutes POST | base=" + __0 +
                          " | server=" + __1 +
                          " | routes=" + FormatNestedRoutes(__result));
            }
            catch
            {
            }
        }

        private static void NetworkMap_FindPhysicalPath_Prefix(string __0, string __1)
        {
            try
            {
                if (!IsInterestingRouteRequest(__0, __1))
                    return;

                LogOnce("FindPhysicalPathPRE|" + __0 + "|" + __1,
                    "ROUTE PATH | FindPhysicalPath PRE | start=" + __0 + " | target=" + __1);
            }
            catch
            {
            }
        }

        private static void NetworkMap_FindPhysicalPath_Postfix(string __0, string __1, object __result)
        {
            try
            {
                if (!IsInterestingRouteRequest(__0, __1))
                    return;

                LogToFile("ROUTE PATH | FindPhysicalPath POST | start=" + __0 +
                          " | target=" + __1 +
                          " | path=" + FormatStringList(__result));
            }
            catch
            {
            }
        }

        private static void NetworkMap_UpdateCustomerServerCountAndSpeed_Prefix(int __0, int __1, float __2)
        {
            try
            {
                LogOnce("UpdateCustomerPRE|" + __0.ToString(CultureInfo.InvariantCulture) + "|" + __1.ToString(CultureInfo.InvariantCulture) + "|" + __2.ToString("0.##", CultureInfo.InvariantCulture),
                    "THROUGHPUT | UpdateCustomerServerCountAndSpeed PRE | customerId=" + __0.ToString(CultureInfo.InvariantCulture) +
                    " | serverCount=" + __1.ToString(CultureInfo.InvariantCulture) +
                    " | speed=" + __2.ToString("0.##", CultureInfo.InvariantCulture));
            }
            catch
            {
            }
        }

        private static void NetworkMap_UpdateCustomerServerCountAndSpeed_Postfix(int __0, int __1, float __2)
        {
            try
            {
                LogToFile("THROUGHPUT | UpdateCustomerServerCountAndSpeed POST | customerId=" + __0.ToString(CultureInfo.InvariantCulture) +
                          " | serverCount=" + __1.ToString(CultureInfo.InvariantCulture) +
                          " | speed=" + __2.ToString("0.##", CultureInfo.InvariantCulture));
            }
            catch
            {
            }
        }

        private static void NetworkMap_GetLACPGroupForCable_Postfix(int __0, object __result)
        {
            try
            {
                if (!WatchedCableIds.Contains(__0))
                    return;

                LogToFile("LACP QUERY | GetLACPGroupForCable | cableId=" + __0.ToString(CultureInfo.InvariantCulture) +
                          " | result=" + FormatLacpGroup(__result));
            }
            catch
            {
            }
        }

        private static void NetworkMap_GetLACPGroupBetween_Postfix(string __0, string __1, object __result)
        {
            try
            {
                if (!StringLooksWatchedDevice(__0) && !StringLooksWatchedDevice(__1))
                    return;

                LogToFile("LACP QUERY | GetLACPGroupBetween | deviceA=" + __0 +
                          " | deviceB=" + __1 +
                          " | result=" + FormatLacpGroup(__result));
            }
            catch
            {
            }
        }

        private void DetectRemoteDeviceBundles(List<List<RegisteredSwitchInfo>> fabrics)
        {
            FabricIdToSpeed.Clear();
            FabricIdToBundleCableIds.Clear();
            WatchedCableIds.Clear();
            WatchedDevices.Clear();

            int index = 1;

            foreach (List<RegisteredSwitchInfo> fabric in fabrics)
            {
                string fabricId = "FABRIC-" + index.ToString("000", CultureInfo.InvariantCulture);
                float fabricSpeed = fabric.Sum(x => x.EstimatedCapacityGbps);
                FabricIdToSpeed[fabricId] = fabricSpeed;
                FabricIdToBundleCableIds[fabricId] = new HashSet<int>();

                HashSet<string> fabricDevices = new HashSet<string>(
                    fabric.Select(x => x.DeviceName).Where(x => !string.IsNullOrWhiteSpace(x)),
                    StringComparer.OrdinalIgnoreCase);

                Dictionary<string, List<RegisteredCableInfo>> remoteBundles =
                    new Dictionary<string, List<RegisteredCableInfo>>(StringComparer.OrdinalIgnoreCase);

                foreach (RegisteredCableInfo cable in RegisteredCables.Values)
                {
                    bool aIn = fabricDevices.Contains(cable.DeviceA);
                    bool bIn = fabricDevices.Contains(cable.DeviceB);

                    if (aIn && bIn)
                    {
                        FabricIdToBundleCableIds[fabricId].Add(cable.CableId);
                        if (!string.IsNullOrWhiteSpace(cable.DeviceA)) WatchedDevices.Add(cable.DeviceA);
                        if (!string.IsNullOrWhiteSpace(cable.DeviceB)) WatchedDevices.Add(cable.DeviceB);
                        continue;
                    }

                    if (aIn == bIn)
                        continue;

                    string remoteKey = aIn
                        ? BuildExactRemoteDeviceKey(cable, true)
                        : BuildExactRemoteDeviceKey(cable, false);

                    if (string.IsNullOrWhiteSpace(remoteKey))
                        continue;

                    List<RegisteredCableInfo> bundle;
                    if (!remoteBundles.TryGetValue(remoteKey, out bundle))
                    {
                        bundle = new List<RegisteredCableInfo>();
                        remoteBundles[remoteKey] = bundle;
                    }

                    bundle.Add(cable);
                }

                foreach (KeyValuePair<string, List<RegisteredCableInfo>> kvp in remoteBundles)
                {
                    List<RegisteredCableInfo> bundle = kvp.Value;
                    if (bundle.Count < 2)
                        continue;

                    foreach (RegisteredCableInfo cable in bundle)
                    {
                        FabricIdToBundleCableIds[fabricId].Add(cable.CableId);
                        WatchedCableIds.Add(cable.CableId);

                        if (!string.IsNullOrWhiteSpace(cable.DeviceA)) WatchedDevices.Add(cable.DeviceA);
                        if (!string.IsNullOrWhiteSpace(cable.DeviceB)) WatchedDevices.Add(cable.DeviceB);
                    }

                    string cableList = string.Join(",", bundle.Select(x => x.CableId).OrderBy(x => x));
                    string logKey = fabricId + "|" + kvp.Key + "|" + bundle.Count.ToString(CultureInfo.InvariantCulture) + "|" + cableList;
                    if (LoggedBundles.Add(logKey))
                    {
                        LogToFile("REMOTE DEVICE BUNDLE | fabricId=" + fabricId +
                                  " | remoteKey=" + kvp.Key +
                                  " | cableCount=" + bundle.Count.ToString(CultureInfo.InvariantCulture) +
                                  " | estPerCableGbps=" + (fabricSpeed / bundle.Count).ToString("0.##", CultureInfo.InvariantCulture) +
                                  " | cableIds=[" + cableList + "]");
                    }
                }

                index++;
            }
        }

        private static bool IsInterestingRouteRequest(string a, string b)
        {
            if (StringLooksWatchedDevice(a) || StringLooksWatchedDevice(b))
                return true;

            if (!string.IsNullOrWhiteSpace(a) && a.StartsWith("Server.", StringComparison.OrdinalIgnoreCase))
                return true;

            if (!string.IsNullOrWhiteSpace(b) && b.StartsWith("Server.", StringComparison.OrdinalIgnoreCase))
                return true;

            return false;
        }

        private static bool StringLooksWatchedDevice(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return false;

            if (WatchedDevices.Contains(value))
                return true;

            foreach (RegisteredCableInfo cable in RegisteredCables.Values)
            {
                if (!WatchedCableIds.Contains(cable.CableId))
                    continue;

                if (string.Equals(cable.DeviceA, value, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(cable.DeviceB, value, StringComparison.OrdinalIgnoreCase))
                    return true;
            }

            return false;
        }

        private static void LogOnce(string key, string message)
        {
            if (LoggedProbeEvents.Add(key))
                LogToFile(message);
        }

        private static string FormatNestedRoutes(object raw)
        {
            if (raw == null)
                return "<null>";

            try
            {
                IEnumerable outer = raw as IEnumerable;
                if (outer == null || raw is string)
                    return raw.ToString();

                List<string> groups = new List<string>();
                int outerCount = 0;

                foreach (object inner in outer)
                {
                    if (outerCount >= 8)
                    {
                        groups.Add("...");
                        break;
                    }

                    groups.Add("[" + FormatStringList(inner) + "]");
                    outerCount++;
                }

                return string.Join(" | ", groups);
            }
            catch
            {
                return "<err>";
            }
        }

        private static string FormatStringList(object raw)
        {
            if (raw == null)
                return "<null>";

            try
            {
                IEnumerable enumerable = raw as IEnumerable;
                if (enumerable == null || raw is string)
                    return raw.ToString();

                List<string> items = new List<string>();
                int count = 0;

                foreach (object item in enumerable)
                {
                    if (count >= 20)
                    {
                        items.Add("...");
                        break;
                    }

                    items.Add(item == null ? "<null>" : item.ToString());
                    count++;
                }

                return string.Join(" -> ", items);
            }
            catch
            {
                return "<err>";
            }
        }

        private static string FormatLacpGroup(object groupObj)
        {
            if (groupObj == null)
                return "<null>";

            try
            {
                object groupId = TryGetMemberValue(groupObj, "groupId");
                object deviceA = TryGetMemberValue(groupObj, "deviceA");
                object deviceB = TryGetMemberValue(groupObj, "deviceB");
                List<int> cableIds = ExtractIntList(TryGetMemberValue(groupObj, "cableIds"));

                return "groupId=" + SafeString(groupId) +
                       ",deviceA=" + SafeString(deviceA) +
                       ",deviceB=" + SafeString(deviceB) +
                       ",cableIds=[" + string.Join(",", cableIds.OrderBy(x => x)) + "]";
            }
            catch
            {
                return groupObj.GetType().Name;
            }
        }

        private static string SafeString(object value)
        {
            return value == null ? "<null>" : Convert.ToString(value, CultureInfo.InvariantCulture);
        }

        private static string BuildExactRemoteDeviceKey(RegisteredCableInfo cable, bool sideAIsLocal)
        {
            string remoteDevice = sideAIsLocal ? cable.DeviceB : cable.DeviceA;
            string remoteExtra = sideAIsLocal ? cable.ExtraB : cable.ExtraA;
            string remoteServerRoot = sideAIsLocal ? cable.ServerRootKeyB : cable.ServerRootKeyA;
            object remoteEndpointObj = sideAIsLocal ? cable.EndpointObjectB : cable.EndpointObjectA;

            if (!string.IsNullOrWhiteSpace(remoteDevice) &&
                remoteDevice.StartsWith("Switch", StringComparison.OrdinalIgnoreCase))
            {
                return "REMOTE_SWITCH_DEVICE:" + remoteDevice.Trim();
            }

            if (!string.IsNullOrWhiteSpace(remoteServerRoot) &&
                remoteServerRoot.StartsWith("SERVERROOT:", StringComparison.OrdinalIgnoreCase))
            {
                return "REMOTE_SERVER_ROOT:" + remoteServerRoot;
            }

            string endpointRoot = ResolveEndpointRootKey(remoteEndpointObj, remoteExtra, remoteDevice);
            if (!string.IsNullOrWhiteSpace(endpointRoot))
            {
                if (endpointRoot.StartsWith("SWITCHROOT:", StringComparison.OrdinalIgnoreCase))
                    return "REMOTE_SWITCH_ROOT:" + endpointRoot;

                if (endpointRoot.StartsWith("SERVERROOT:", StringComparison.OrdinalIgnoreCase))
                    return "REMOTE_SERVER_ROOT:" + endpointRoot;
            }

            if (!string.IsNullOrWhiteSpace(remoteServerRoot) &&
                remoteServerRoot.StartsWith("SERVERFAMILY:", StringComparison.OrdinalIgnoreCase))
            {
                return "REMOTE_SERVER_FAMILY:" + remoteServerRoot;
            }

            if (!string.IsNullOrWhiteSpace(remoteExtra) &&
                remoteExtra.StartsWith("Server.", StringComparison.OrdinalIgnoreCase))
            {
                return "REMOTE_SERVER_FAMILY:" + NormalizeRuntimeIdentity(remoteExtra);
            }

            if (!string.IsNullOrWhiteSpace(remoteDevice) &&
                remoteDevice.StartsWith("Server.", StringComparison.OrdinalIgnoreCase))
            {
                return "REMOTE_SERVER_FAMILY:" + NormalizeRuntimeIdentity(remoteDevice);
            }

            return string.Empty;
        }

        private static RegisteredSwitchInfo BuildRegisteredSwitchInfo(object networkSwitchObj)
        {
            if (networkSwitchObj == null)
                return null;

            string deviceName = DiscoverBestDeviceName(networkSwitchObj);
            string modelName = DiscoverBestModelName(networkSwitchObj);

            GameObject go = DiscoverGameObject(networkSwitchObj);
            Vector3 pos = go != null ? go.transform.position : Vector3.zero;

            RegisteredSwitchInfo info = new RegisteredSwitchInfo();
            info.NetworkSwitchObject = networkSwitchObj;
            info.GameObject = go;
            info.DeviceName = deviceName;
            info.ModelName = modelName;
            info.WorldPosition = pos;
            info.HasWorldPosition = go != null;

            return info;
        }

        private void RefreshRegisteredSwitchPositions()
        {
            foreach (RegisteredSwitchInfo info in RegisteredSwitches.Values)
            {
                if (info == null)
                    continue;

                if (info.GameObject == null && info.NetworkSwitchObject != null)
                    info.GameObject = DiscoverGameObject(info.NetworkSwitchObject);

                if (info.GameObject != null)
                {
                    info.WorldPosition = info.GameObject.transform.position;
                    info.HasWorldPosition = true;
                }

                if (string.IsNullOrWhiteSpace(info.ModelName))
                    info.ModelName = DiscoverBestModelName(info.NetworkSwitchObject);

                info.EstimatedCapacityGbps = EstimateCapacityFromModel(info.ModelName);
            }
        }

        private static bool IsAllowedSwitchModel(string modelName)
        {
            if (string.IsNullOrWhiteSpace(modelName))
                return false;

            string normalized = NormalizeCloneName(modelName);
            return AllowedSwitchModels.Contains(normalized);
        }

        private static bool IsLikelyPlacedSwitch(RegisteredSwitchInfo info)
        {
            if (info == null || !info.HasWorldPosition)
                return false;

            float xzRadius = Mathf.Sqrt((info.WorldPosition.x * info.WorldPosition.x) + (info.WorldPosition.z * info.WorldPosition.z));
            if (xzRadius <= OriginXZRejectRadius)
                return false;

            if (info.WorldPosition.y < MinUsefulWorldY || info.WorldPosition.y > MaxUsefulWorldY)
                return false;

            return true;
        }

        private static float EstimateCapacityFromModel(string modelName)
        {
            string normalized = NormalizeCloneName(modelName);

            if (normalized.Equals("Switch32xQSFP", StringComparison.OrdinalIgnoreCase))
                return 32f * 40f;

            if (normalized.Equals("Switch4xQSXP16xSFP", StringComparison.OrdinalIgnoreCase))
                return (4f * 40f) + (16f * 10f);

            if (normalized.Equals("Switch4xSFP", StringComparison.OrdinalIgnoreCase))
                return 4f * 10f;

            if (normalized.Equals("Switch16CU", StringComparison.OrdinalIgnoreCase))
                return 16f * 1f;

            return 0f;
        }

        private static List<List<RegisteredSwitchInfo>> BuildRegistryFabrics(List<RegisteredSwitchInfo> switches)
        {
            List<List<RegisteredSwitchInfo>> groups = new List<List<RegisteredSwitchInfo>>();
            HashSet<RegisteredSwitchInfo> remaining = new HashSet<RegisteredSwitchInfo>(switches);

            while (remaining.Count > 0)
            {
                RegisteredSwitchInfo seed = remaining.First();
                remaining.Remove(seed);

                List<RegisteredSwitchInfo> group = new List<RegisteredSwitchInfo>();
                Queue<RegisteredSwitchInfo> queue = new Queue<RegisteredSwitchInfo>();
                queue.Enqueue(seed);

                while (queue.Count > 0)
                {
                    RegisteredSwitchInfo current = queue.Dequeue();
                    group.Add(current);

                    List<RegisteredSwitchInfo> neighbors = remaining
                        .Where(other => AreAdjacent(current, other))
                        .ToList();

                    foreach (RegisteredSwitchInfo neighbor in neighbors)
                    {
                        remaining.Remove(neighbor);
                        queue.Enqueue(neighbor);
                    }
                }

                groups.Add(group);
            }

            return groups;
        }

        private static bool AreAdjacent(RegisteredSwitchInfo a, RegisteredSwitchInfo b)
        {
            if (a == null || b == null || !a.HasWorldPosition || !b.HasWorldPosition)
                return false;

            float dx = Mathf.Abs(a.WorldPosition.x - b.WorldPosition.x);
            float dz = Mathf.Abs(a.WorldPosition.z - b.WorldPosition.z);

            if (dx > SameFabricXZTolerance || dz > SameFabricXZTolerance)
                return false;

            float actualStep = Mathf.Abs(a.WorldPosition.y - b.WorldPosition.y);
            float expectedStep = FallbackExpectedStep;

            return Mathf.Abs(actualStep - expectedStep) <= AdjacentYTolerance;
        }

        private static string ResolveEndpointRootKey(object endpointObj, string extraName, string deviceName)
        {
            GameObject endpointGo = DiscoverGameObject(endpointObj);
            if (endpointGo != null)
            {
                Transform current = endpointGo.transform;
                int depth = 0;

                while (current != null && depth < 16)
                {
                    string n = NormalizeCloneName(current.name);

                    if (n.StartsWith("Server.", StringComparison.OrdinalIgnoreCase))
                        return "SERVERROOT:" + n + "#" + current.gameObject.GetInstanceID().ToString(CultureInfo.InvariantCulture);

                    if (n.StartsWith("Switch", StringComparison.OrdinalIgnoreCase))
                        return "SWITCHROOT:" + n + "#" + current.gameObject.GetInstanceID().ToString(CultureInfo.InvariantCulture);

                    current = current.parent;
                    depth++;
                }
            }

            string raw = !string.IsNullOrWhiteSpace(extraName) ? extraName : deviceName;
            if (!string.IsNullOrWhiteSpace(raw))
            {
                string normalized = NormalizeRuntimeIdentity(raw);

                if (normalized.StartsWith("Server.", StringComparison.OrdinalIgnoreCase))
                    return "SERVERFAMILY:" + normalized;

                if (normalized.StartsWith("Switch", StringComparison.OrdinalIgnoreCase))
                    return "SWITCHFAMILY:" + normalized;
            }

            return string.Empty;
        }

        private static string ResolveServerRootKey(object endpointObj, string extraName, string deviceName)
        {
            GameObject endpointGo = DiscoverGameObject(endpointObj);
            if (endpointGo != null)
            {
                Transform current = endpointGo.transform;
                int depth = 0;

                while (current != null && depth < 16)
                {
                    string n = NormalizeCloneName(current.name);
                    if (n.StartsWith("Server.", StringComparison.OrdinalIgnoreCase))
                        return "SERVERROOT:" + n + "#" + current.gameObject.GetInstanceID().ToString(CultureInfo.InvariantCulture);

                    current = current.parent;
                    depth++;
                }
            }

            string raw = !string.IsNullOrWhiteSpace(extraName) ? extraName : deviceName;
            if (!string.IsNullOrWhiteSpace(raw) && raw.StartsWith("Server.", StringComparison.OrdinalIgnoreCase))
            {
                string normalized = NormalizeRuntimeIdentity(raw);
                return "SERVERFAMILY:" + normalized;
            }

            return string.Empty;
        }

        private static GameObject DiscoverGameObject(object obj)
        {
            if (obj == null)
                return null;

            try
            {
                GameObject go = obj as GameObject;
                if (go != null)
                    return go;
            }
            catch
            {
            }

            try
            {
                Component comp = obj as Component;
                if (comp != null)
                    return comp.gameObject;
            }
            catch
            {
            }

            try
            {
                object go = TryGetMemberValue(obj, "gameObject");
                GameObject gameObject = go as GameObject;
                if (gameObject != null)
                    return gameObject;
            }
            catch
            {
            }

            try
            {
                object transform = TryGetMemberValue(obj, "transform");
                Transform t = transform as Transform;
                if (t != null)
                    return t.gameObject;
            }
            catch
            {
            }

            return null;
        }

        private static string DiscoverBestDeviceName(object obj)
        {
            string[] candidates =
            {
                "deviceName", "DeviceName", "switchId", "SwitchId", "networkId", "NetworkId", "id", "Id"
            };

            foreach (string name in candidates)
            {
                object val = TryGetMemberValue(obj, name);
                string s = val as string;
                if (!string.IsNullOrWhiteSpace(s))
                    return s.Trim();
            }

            foreach (FieldInfo f in SafeFields(obj.GetType()))
            {
                if (f.FieldType != typeof(string))
                    continue;

                string lower = f.Name.ToLowerInvariant();
                if (!(lower.Contains("device") || lower.Contains("switch") || lower.Contains("network") || lower == "id"))
                    continue;

                try
                {
                    string s = f.GetValue(obj) as string;
                    if (!string.IsNullOrWhiteSpace(s))
                        return s.Trim();
                }
                catch
                {
                }
            }

            foreach (PropertyInfo p in SafeProperties(obj.GetType()))
            {
                if (p.PropertyType != typeof(string) || !p.CanRead)
                    continue;

                string lower = p.Name.ToLowerInvariant();
                if (!(lower.Contains("device") || lower.Contains("switch") || lower.Contains("network") || lower == "id"))
                    continue;

                try
                {
                    string s = p.GetValue(obj, null) as string;
                    if (!string.IsNullOrWhiteSpace(s))
                        return s.Trim();
                }
                catch
                {
                }
            }

            return string.Empty;
        }

        private static string DiscoverBestModelName(object obj)
        {
            GameObject go = DiscoverGameObject(obj);
            if (go != null)
            {
                string normalized = NormalizeCloneName(go.name);
                if (AllowedSwitchModels.Contains(normalized))
                    return normalized;
            }

            foreach (string name in AllowedSwitchModels)
            {
                if (go != null && go.name.IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0)
                    return name;
            }

            return go != null ? NormalizeCloneName(go.name) : string.Empty;
        }

        private static IEnumerable<FieldInfo> SafeFields(Type t)
        {
            try { return t.GetFields(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance); }
            catch { return new FieldInfo[0]; }
        }

        private static IEnumerable<PropertyInfo> SafeProperties(Type t)
        {
            try { return t.GetProperties(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance); }
            catch { return new PropertyInfo[0]; }
        }

        private static object TryGetMemberValue(object target, string memberName)
        {
            if (target == null || string.IsNullOrWhiteSpace(memberName))
                return null;

            try
            {
                Type t = target.GetType();

                PropertyInfo prop = t.GetProperty(memberName, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                if (prop != null && prop.CanRead)
                    return prop.GetValue(target, null);

                FieldInfo field = t.GetField(memberName, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                if (field != null)
                    return field.GetValue(target);
            }
            catch
            {
            }

            return null;
        }

        private static List<int> ExtractIntList(object raw)
        {
            List<int> result = new List<int>();

            if (raw == null)
                return result;

            try
            {
                if (raw is int)
                {
                    result.Add((int)raw);
                    return result;
                }

                IEnumerable enumerable = raw as IEnumerable;
                if (enumerable != null && !(raw is string))
                {
                    foreach (object item in enumerable)
                    {
                        if (item == null)
                            continue;

                        int parsed;
                        if (int.TryParse(item.ToString(), out parsed))
                            result.Add(parsed);
                    }

                    return result.Distinct().ToList();
                }

                int single;
                if (int.TryParse(raw.ToString(), out single))
                    result.Add(single);
            }
            catch
            {
            }

            return result.Distinct().ToList();
        }

        private static string NormalizeCloneName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return string.Empty;

            string result = name.Trim();
            if (result.EndsWith("(Clone)", StringComparison.OrdinalIgnoreCase))
                result = result.Substring(0, result.Length - "(Clone)".Length).Trim();

            return result;
        }

        private static string NormalizeRuntimeIdentity(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
                return string.Empty;

            return TrailingRuntimeIdRegex.Replace(raw.Trim(), "");
        }

        private static void LogToFile(string message)
        {
            try
            {
                Directory.CreateDirectory(DebugFolderPath);
                File.AppendAllText(DebugLogPath, "[" + DateTime.Now.ToString("HH:mm:ss.fff") + "] " + message + Environment.NewLine);
            }
            catch
            {
            }
        }
    }

    public sealed class FabricGroupTag : MonoBehaviour
    {
        public FabricGroupTag(IntPtr ptr) : base(ptr) { }

        public string FabricId { get; set; }
        public int MemberIndex { get; set; }
        public int MemberCount { get; set; }
        public float AggregatedBandwidth { get; set; }
        public int AggregatedPortCount { get; set; }
    }

    internal sealed class RegisteredSwitchInfo
    {
        public object NetworkSwitchObject;
        public GameObject GameObject;
        public string DeviceName;
        public string ModelName;
        public Vector3 WorldPosition;
        public bool HasWorldPosition;
        public float EstimatedCapacityGbps;
    }

    internal sealed class RegisteredCableInfo
    {
        public int CableId;
        public string DeviceA;
        public string DeviceB;
        public string ExtraA;
        public string ExtraB;
        public int PortA;
        public int PortB;

        public object EndpointObjectA;
        public object EndpointObjectB;

        public string ServerRootKeyA;
        public string ServerRootKeyB;
    }
}