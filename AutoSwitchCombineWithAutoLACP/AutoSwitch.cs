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
using System.Text;
using System.Text.RegularExpressions;
using UnityEngine;
using Component = UnityEngine.Component;
using Object = UnityEngine.Object;

[assembly: MelonInfo(typeof(AutoSwitch.AutoSwitchMod), "AutoSwitch", "2.7.0", "Big Texas Jerky")]
[assembly: MelonGame("Waseku", "Data Center")]

namespace AutoSwitch
{
    public sealed class AutoSwitchMod : MelonMod
    {
        private const float ScanIntervalSeconds = 6.0f;

        private const float SameFabricXZTolerance = 0.40f;
        private const float AdjacentYTolerance = 0.040f;
        private const float FallbackExpectedStep = 0.0444f;

        private const float OriginXZRejectRadius = 1.25f;
        private const float MinUsefulWorldY = 1.0f;
        private const float MaxUsefulWorldY = 4.5f;

        private static readonly Regex TrailingRuntimeIdRegex =
            new Regex(@"_[\-]?\d+$", RegexOptions.Compiled);

        private static readonly string[] ProbeKeywords =
        {
            "active", "selected", "preferred", "primary",
            "route", "path", "traffic", "load", "usage",
            "bandwidth", "speed", "throughput", "connected",
            "source", "destination", "dest", "from", "to",
            "endpoint", "port", "cable", "link"
        };

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

        private static readonly Dictionary<string, HashSet<int>> FabricIdToCandidateCableIds =
            new Dictionary<string, HashSet<int>>(StringComparer.OrdinalIgnoreCase);

        private static readonly HashSet<int> WatchedCableIds =
            new HashSet<int>();

        private static readonly HashSet<string> LoggedPatches =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private static readonly HashSet<string> LoggedBundles =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private static readonly HashSet<string> LoggedCableLinkSignatures =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private static readonly HashSet<string> LoggedCableLinkTypes =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private static readonly Dictionary<int, float> LastSetConnectionSpeed =
            new Dictionary<int, float>();

        private static Type _networkMapType;
        private static Type _cableLinkType;
        private static Type _lacpGroupType;

        private float _nextScanTime;
        private string _lastSummary = string.Empty;

        public override void OnInitializeMelon()
        {
            ClassInjector.RegisterTypeInIl2Cpp<FabricGroupTag>();

            Directory.CreateDirectory(DebugFolderPath);
            File.WriteAllText(DebugLogPath, "[" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "] AutoSwitch 2.7 debug log started." + Environment.NewLine);

            InstallNativePatches();

            MelonLogger.Msg("[AutoSwitch] v2.7.0 active. CableLink probe build.");
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
            FabricIdToCandidateCableIds.Clear();
            WatchedCableIds.Clear();
            LoggedBundles.Clear();
            LoggedCableLinkSignatures.Clear();
            LoggedCableLinkTypes.Clear();
            LastSetConnectionSpeed.Clear();

            LogToFile("Scene loaded: " + sceneName + " (" + buildIndex.ToString(CultureInfo.InvariantCulture) + ")");
        }

        public override void OnUpdate()
        {
            try
            {
                if (Time.time < _nextScanTime)
                    return;

                _nextScanTime = Time.time + ScanIntervalSeconds;
                RunProbePass();
            }
            catch (Exception ex)
            {
                MelonLogger.Error("[AutoSwitch] OnUpdate exception: " + ex);
                LogToFile("OnUpdate exception: " + ex);
                _nextScanTime = Time.time + 10.0f;
            }
        }

        private void RunProbePass()
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
                " | watchedCableIds=" + WatchedCableIds.Count.ToString(CultureInfo.InvariantCulture);

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
                    int tracked = FabricIdToCandidateCableIds.ContainsKey(fabricId) ? FabricIdToCandidateCableIds[fabricId].Count : 0;

                    string members = string.Join(", ",
                        fabric.Select(x => x.DeviceName + "@(" +
                                           x.WorldPosition.x.ToString("0.###", CultureInfo.InvariantCulture) + "," +
                                           x.WorldPosition.y.ToString("0.###", CultureInfo.InvariantCulture) + "," +
                                           x.WorldPosition.z.ToString("0.###", CultureInfo.InvariantCulture) + ")[" +
                                           x.ModelName + "]"));

                    LogToFile("FABRIC | id=" + fabricId +
                              " | members=" + fabric.Count.ToString(CultureInfo.InvariantCulture) +
                              " | estCapacityGbps=" + speed.ToString("0.##", CultureInfo.InvariantCulture) +
                              " | candidateCableIds=" + tracked.ToString(CultureInfo.InvariantCulture) +
                              " | switches=" + members);

                    if (tracked > 0)
                    {
                        string cableList = string.Join(",", FabricIdToCandidateCableIds[fabricId].OrderBy(x => x));
                        LogToFile("FABRIC CANDIDATE CABLE MAP | fabricId=" + fabricId + " | cableIds=[" + cableList + "]");
                    }

                    fabricIndex++;
                }
            }

            ProbeCableLinks();
        }

        private void InstallNativePatches()
        {
            try
            {
                _networkMapType = AccessTools.TypeByName("Il2Cpp.NetworkMap");
                _cableLinkType = AccessTools.TypeByName("Il2Cpp.CableLink");
                _lacpGroupType = AccessTools.TypeByName("Il2Cpp.NetworkMap+LACPGroup");

                if (_networkMapType != null)
                {
                    MethodInfo registerSwitch = AccessTools.Method(_networkMapType, "RegisterSwitch");
                    if (registerSwitch != null)
                    {
                        HarmonyInstance.Patch(
                            registerSwitch,
                            postfix: new HarmonyMethod(typeof(AutoSwitchMod).GetMethod(nameof(NetworkMap_RegisterSwitch_Postfix), BindingFlags.NonPublic | BindingFlags.Static))
                        );
                        LogPatch("Patched NetworkMap.RegisterSwitch");
                    }

                    MethodInfo registerCable = AccessTools.Method(_networkMapType, "RegisterCableConnection");
                    if (registerCable != null)
                    {
                        HarmonyInstance.Patch(
                            registerCable,
                            postfix: new HarmonyMethod(typeof(AutoSwitchMod).GetMethod(nameof(NetworkMap_RegisterCableConnection_Postfix), BindingFlags.NonPublic | BindingFlags.Static))
                        );
                        LogPatch("Patched NetworkMap.RegisterCableConnection");
                    }

                    MethodInfo createLacp = AccessTools.Method(_networkMapType, "CreateLACPGroup");
                    if (createLacp != null)
                    {
                        HarmonyInstance.Patch(
                            createLacp,
                            postfix: new HarmonyMethod(typeof(AutoSwitchMod).GetMethod(nameof(NetworkMap_CreateLACPGroup_Postfix), BindingFlags.NonPublic | BindingFlags.Static))
                        );
                        LogPatch("Patched NetworkMap.CreateLACPGroup");
                    }
                }

                if (_lacpGroupType != null)
                {
                    MethodInfo aggSpeed = AccessTools.Method(_lacpGroupType, "GetAggregatedSpeed");
                    if (aggSpeed != null)
                    {
                        HarmonyInstance.Patch(
                            aggSpeed,
                            postfix: new HarmonyMethod(typeof(AutoSwitchMod).GetMethod(nameof(LACPGroup_GetAggregatedSpeed_Postfix), BindingFlags.NonPublic | BindingFlags.Static))
                        );
                        LogPatch("Patched NetworkMap+LACPGroup.GetAggregatedSpeed");
                    }
                }

                if (_cableLinkType != null)
                {
                    MethodInfo setConnectionSpeed = AccessTools.Method(_cableLinkType, "SetConnectionSpeed", new Type[] { typeof(float) });
                    if (setConnectionSpeed != null)
                    {
                        HarmonyInstance.Patch(
                            setConnectionSpeed,
                            prefix: new HarmonyMethod(typeof(AutoSwitchMod).GetMethod(nameof(CableLink_SetConnectionSpeed_Prefix), BindingFlags.NonPublic | BindingFlags.Static)),
                            postfix: new HarmonyMethod(typeof(AutoSwitchMod).GetMethod(nameof(CableLink_SetConnectionSpeed_Postfix), BindingFlags.NonPublic | BindingFlags.Static))
                        );
                        LogPatch("Patched CableLink.SetConnectionSpeed");
                    }
                }
            }
            catch (Exception ex)
            {
                LogToFile("InstallNativePatches failed: " + ex);
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

        private static void NetworkMap_CreateLACPGroup_Postfix(string __0, string __1, object __2, ref int __result)
        {
            try
            {
                List<int> cableIds = ExtractIntList(__2);
                string joined = string.Join(",", cableIds.OrderBy(x => x));

                LogToFile("CREATE LACP | groupId=" + __result.ToString(CultureInfo.InvariantCulture) +
                          " | deviceA=" + __0 +
                          " | deviceB=" + __1 +
                          " | cableIds=[" + joined + "]");
            }
            catch (Exception ex)
            {
                LogToFile("NetworkMap_CreateLACPGroup_Postfix failed: " + ex);
            }
        }

        private static void LACPGroup_GetAggregatedSpeed_Postfix(object __instance, ref float __result)
        {
            try
            {
                if (__instance == null)
                    return;

                object raw = TryGetMemberValue(__instance, "cableIds");
                List<int> cableIds = ExtractIntList(raw);

                if (cableIds.Any(x => WatchedCableIds.Contains(x)))
                {
                    string ids = string.Join(",", cableIds.OrderBy(x => x));
                    LogToFile("LACP AGG SPEED | cableIds=[" + ids + "] | result=" + __result.ToString("0.##", CultureInfo.InvariantCulture));
                }
            }
            catch
            {
            }
        }

        private static void CableLink_SetConnectionSpeed_Prefix(object __instance, ref float __0)
        {
            try
            {
                if (__instance == null)
                    return;

                List<int> cableIds = ExtractIntList(TryGetMemberValue(__instance, "cableIDsOnLink"));
                if (cableIds.Count == 0)
                    return;

                if (!cableIds.Any(x => WatchedCableIds.Contains(x)))
                    return;

                int key = GetStableObjectKey(__instance);
                LastSetConnectionSpeed[key] = __0;

                LogToFile("CABLELINK SET SPEED PRE | key=" + key.ToString(CultureInfo.InvariantCulture) +
                          " | requested=" + __0.ToString("0.##", CultureInfo.InvariantCulture) +
                          " | cableIds=[" + string.Join(",", cableIds.OrderBy(x => x)) + "]");
            }
            catch
            {
            }
        }

        private static void CableLink_SetConnectionSpeed_Postfix(object __instance, float __0)
        {
            try
            {
                if (__instance == null)
                    return;

                List<int> cableIds = ExtractIntList(TryGetMemberValue(__instance, "cableIDsOnLink"));
                if (cableIds.Count == 0)
                    return;

                if (!cableIds.Any(x => WatchedCableIds.Contains(x)))
                    return;

                int key = GetStableObjectKey(__instance);
                float previous = LastSetConnectionSpeed.ContainsKey(key) ? LastSetConnectionSpeed[key] : __0;

                LogToFile("CABLELINK SET SPEED POST | key=" + key.ToString(CultureInfo.InvariantCulture) +
                          " | requested=" + __0.ToString("0.##", CultureInfo.InvariantCulture) +
                          " | previousSeen=" + previous.ToString("0.##", CultureInfo.InvariantCulture) +
                          " | cableIds=[" + string.Join(",", cableIds.OrderBy(x => x)) + "]");
            }
            catch
            {
            }
        }

        private void DetectRemoteDeviceBundles(List<List<RegisteredSwitchInfo>> fabrics)
        {
            FabricIdToSpeed.Clear();
            FabricIdToCandidateCableIds.Clear();
            WatchedCableIds.Clear();

            int index = 1;

            foreach (List<RegisteredSwitchInfo> fabric in fabrics)
            {
                string fabricId = "FABRIC-" + index.ToString("000", CultureInfo.InvariantCulture);
                float fabricSpeed = fabric.Sum(x => x.EstimatedCapacityGbps);
                FabricIdToSpeed[fabricId] = fabricSpeed;
                FabricIdToCandidateCableIds[fabricId] = new HashSet<int>();

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
                        FabricIdToCandidateCableIds[fabricId].Add(cable.CableId);
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
                        FabricIdToCandidateCableIds[fabricId].Add(cable.CableId);
                        WatchedCableIds.Add(cable.CableId);
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

        private void ProbeCableLinks()
        {
            try
            {
                if (_cableLinkType == null)
                    return;

                MethodInfo findAllMethod = typeof(Resources).GetMethod("FindObjectsOfTypeAll", new Type[] { typeof(Type) });
                if (findAllMethod == null)
                {
                    LogToFile("PROBE | Resources.FindObjectsOfTypeAll(Type) not found.");
                    return;
                }

                Object[] found = findAllMethod.Invoke(null, new object[] { _cableLinkType }) as Object[];
                if (found == null)
                {
                    LogToFile("PROBE | No CableLink array returned.");
                    return;
                }

                LogCableLinkTypeMetadata(_cableLinkType);

                int watchedMatches = 0;

                foreach (Object obj in found)
                {
                    object cableLinkObj = obj;
                    if (cableLinkObj == null)
                        continue;

                    List<int> cableIds = ExtractIntList(TryGetMemberValue(cableLinkObj, "cableIDsOnLink"));
                    if (cableIds.Count == 0)
                        cableIds = ExtractIntList(TryGetMemberValue(cableLinkObj, "cableIds"));

                    if (cableIds.Count == 0)
                        continue;

                    if (!cableIds.Any(x => WatchedCableIds.Contains(x)))
                        continue;

                    watchedMatches++;

                    string signature = BuildCableLinkSignature(cableLinkObj, cableIds);
                    if (LoggedCableLinkSignatures.Add(signature))
                    {
                        LogToFile(BuildCableLinkDump(cableLinkObj, cableIds));
                    }
                }

                LogToFile("PROBE | CableLink watched matches=" + watchedMatches.ToString(CultureInfo.InvariantCulture) +
                          " | watchedCableIds=[" + string.Join(",", WatchedCableIds.OrderBy(x => x)) + "]");
            }
            catch (Exception ex)
            {
                LogToFile("ProbeCableLinks failed: " + ex);
            }
        }

        private static void LogCableLinkTypeMetadata(Type cableLinkType)
        {
            try
            {
                string typeName = cableLinkType.FullName;
                if (!LoggedCableLinkTypes.Add(typeName))
                    return;

                StringBuilder sb = new StringBuilder();
                sb.Append("CABLELINK TYPE META | type=").Append(typeName);

                List<string> methodNames = SafeMethods(cableLinkType)
                    .Select(m => m.Name)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Where(n => NameMatchesProbeKeyword(n))
                    .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                List<string> fieldNames = SafeFields(cableLinkType)
                    .Select(f => f.Name)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Where(n => NameMatchesProbeKeyword(n))
                    .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                List<string> propNames = SafeProperties(cableLinkType)
                    .Select(p => p.Name)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Where(n => NameMatchesProbeKeyword(n))
                    .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                sb.Append(" | methods=[").Append(string.Join(",", methodNames)).Append("]");
                sb.Append(" | fields=[").Append(string.Join(",", fieldNames)).Append("]");
                sb.Append(" | props=[").Append(string.Join(",", propNames)).Append("]");

                LogToFile(sb.ToString());
            }
            catch (Exception ex)
            {
                LogToFile("LogCableLinkTypeMetadata failed: " + ex);
            }
        }

        private static string BuildCableLinkSignature(object cableLinkObj, List<int> cableIds)
        {
            int key = GetStableObjectKey(cableLinkObj);
            string ids = string.Join(",", cableIds.OrderBy(x => x));

            string selected = ReadSimpleProbeValue(cableLinkObj, "selected");
            string active = ReadSimpleProbeValue(cableLinkObj, "active");
            string speed = ReadSimpleProbeValue(cableLinkObj, "connectionSpeed");

            return "CableLink#" + key.ToString(CultureInfo.InvariantCulture) +
                   "|ids=" + ids +
                   "|selected=" + selected +
                   "|active=" + active +
                   "|speed=" + speed;
        }

        private static string BuildCableLinkDump(object cableLinkObj, List<int> cableIds)
        {
            StringBuilder sb = new StringBuilder();

            int key = GetStableObjectKey(cableLinkObj);
            GameObject go = DiscoverGameObject(cableLinkObj);

            sb.Append("CABLELINK PROBE | key=").Append(key.ToString(CultureInfo.InvariantCulture));
            sb.Append(" | type=").Append(cableLinkObj.GetType().FullName);
            sb.Append(" | cableIds=[").Append(string.Join(",", cableIds.OrderBy(x => x))).Append("]");

            if (go != null)
            {
                sb.Append(" | gameObject=").Append(go.name);
                sb.Append(" | instanceId=").Append(go.GetInstanceID().ToString(CultureInfo.InvariantCulture));
            }

            List<string> memberPairs = new List<string>();

            foreach (FieldInfo field in SafeFields(cableLinkObj.GetType()))
            {
                if (!NameMatchesProbeKeyword(field.Name))
                    continue;

                string value = FormatProbeValue(SafeGetFieldValue(field, cableLinkObj));
                memberPairs.Add("F:" + field.Name + "=" + value);
            }

            foreach (PropertyInfo prop in SafeProperties(cableLinkObj.GetType()))
            {
                if (!prop.CanRead)
                    continue;

                if (!NameMatchesProbeKeyword(prop.Name))
                    continue;

                string value = FormatProbeValue(SafeGetPropertyValue(prop, cableLinkObj));
                memberPairs.Add("P:" + prop.Name + "=" + value);
            }

            sb.Append(" | probeMembers={").Append(string.Join(" ; ", memberPairs.Distinct(StringComparer.OrdinalIgnoreCase))).Append("}");

            object endpointsA = TryGetMemberValue(cableLinkObj, "from");
            object endpointsB = TryGetMemberValue(cableLinkObj, "to");
            object source = TryGetMemberValue(cableLinkObj, "source");
            object destination = TryGetMemberValue(cableLinkObj, "destination");
            object endpoint1 = TryGetMemberValue(cableLinkObj, "endpointA");
            object endpoint2 = TryGetMemberValue(cableLinkObj, "endpointB");
            object currentPath = TryGetMemberValue(cableLinkObj, "path");
            object routeObj = TryGetMemberValue(cableLinkObj, "route");

            sb.Append(" | refs={");
            sb.Append("from=").Append(FormatProbeValue(endpointsA)).Append(" ; ");
            sb.Append("to=").Append(FormatProbeValue(endpointsB)).Append(" ; ");
            sb.Append("source=").Append(FormatProbeValue(source)).Append(" ; ");
            sb.Append("destination=").Append(FormatProbeValue(destination)).Append(" ; ");
            sb.Append("endpointA=").Append(FormatProbeValue(endpoint1)).Append(" ; ");
            sb.Append("endpointB=").Append(FormatProbeValue(endpoint2)).Append(" ; ");
            sb.Append("path=").Append(FormatProbeValue(currentPath)).Append(" ; ");
            sb.Append("route=").Append(FormatProbeValue(routeObj));
            sb.Append("}");

            return sb.ToString();
        }

        private static string ReadSimpleProbeValue(object target, string memberName)
        {
            try
            {
                object value = TryGetMemberValue(target, memberName);
                return FormatProbeValue(value);
            }
            catch
            {
                return "<err>";
            }
        }

        private static bool NameMatchesProbeKeyword(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return false;

            string lower = name.ToLowerInvariant();
            foreach (string keyword in ProbeKeywords)
            {
                if (lower.Contains(keyword))
                    return true;
            }

            return false;
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

        private static int GetStableObjectKey(object obj)
        {
            if (obj == null)
                return 0;

            try
            {
                GameObject go = DiscoverGameObject(obj);
                if (go != null)
                    return go.GetInstanceID();
            }
            catch
            {
            }

            try
            {
                PropertyInfo prop = obj.GetType().GetProperty("Pointer", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                if (prop != null)
                {
                    object value = prop.GetValue(obj, null);
                    if (value != null)
                        return value.GetHashCode();
                }
            }
            catch
            {
            }

            return obj.GetHashCode();
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

        private static IEnumerable<MethodInfo> SafeMethods(Type t)
        {
            try { return t.GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance); }
            catch { return new MethodInfo[0]; }
        }

        private static object SafeGetFieldValue(FieldInfo field, object target)
        {
            try { return field.GetValue(target); }
            catch { return "<err>"; }
        }

        private static object SafeGetPropertyValue(PropertyInfo prop, object target)
        {
            try { return prop.GetValue(target, null); }
            catch { return "<err>"; }
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

        private static string FormatProbeValue(object value)
        {
            if (value == null)
                return "<null>";

            if (value is string)
                return "\"" + value.ToString() + "\"";

            if (value is bool || value is byte || value is sbyte ||
                value is short || value is ushort || value is int || value is uint ||
                value is long || value is ulong || value is float || value is double || value is decimal)
                return Convert.ToString(value, CultureInfo.InvariantCulture);

            if (value is Enum)
                return value.ToString();

            if (value is Vector3)
            {
                Vector3 v = (Vector3)value;
                return "(" + v.x.ToString("0.###", CultureInfo.InvariantCulture) + "," +
                             v.y.ToString("0.###", CultureInfo.InvariantCulture) + "," +
                             v.z.ToString("0.###", CultureInfo.InvariantCulture) + ")";
            }

            GameObject go = DiscoverGameObject(value);
            if (go != null)
                return value.GetType().Name + ":" + go.name + "#" + go.GetInstanceID().ToString(CultureInfo.InvariantCulture);

            IEnumerable ints = value as IEnumerable;
            if (ints != null && !(value is string))
            {
                List<string> items = new List<string>();
                int count = 0;
                foreach (object item in ints)
                {
                    if (count >= 12)
                    {
                        items.Add("...");
                        break;
                    }

                    items.Add(item == null ? "<null>" : item.ToString());
                    count++;
                }

                return "[" + string.Join(",", items) + "]";
            }

            return value.GetType().Name + ":" + value;
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