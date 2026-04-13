using HarmonyLib;
using Il2Cpp;
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

[assembly: MelonInfo(typeof(AutoSwitch.AutoSwitchMod), "AutoSwitch", "2.11.1", "Big Texas Jerky")]
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

        internal static readonly Dictionary<string, RegisteredSwitchInfo> RegisteredSwitches =
            new Dictionary<string, RegisteredSwitchInfo>(StringComparer.OrdinalIgnoreCase);

        internal static readonly Dictionary<int, RegisteredCableInfo> RegisteredCables =
            new Dictionary<int, RegisteredCableInfo>();

        private static readonly Dictionary<string, float> FabricIdToSpeed =
            new Dictionary<string, float>(StringComparer.OrdinalIgnoreCase);

        private static readonly Dictionary<string, List<int>> FabricIdToBundleCableIds =
            new Dictionary<string, List<int>>(StringComparer.OrdinalIgnoreCase);

        private static readonly HashSet<string> LoggedPatches =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private static readonly HashSet<string> LoggedBundleSignatures =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private float _nextScanTime;
        private string _lastSummary = string.Empty;

        public override void OnInitializeMelon()
        {
            ClassInjector.RegisterTypeInIl2Cpp<FabricGroupTag>();

            Directory.CreateDirectory(DebugFolderPath);
            File.WriteAllText(DebugLogPath, "[" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "] AutoSwitch 2.11.1 debug log started." + Environment.NewLine);

            InstallNativePatches();
            SaveDataAutoLACP.StartSafeBootstrap();

            MelonLogger.Msg("[AutoSwitch] v2.11.1 active. Save-data endpoint bundle mode.");
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
            LoggedBundleSignatures.Clear();

            SaveDataAutoLACP.ResetForScene();
            SaveDataAutoLACP.StartSafeBootstrap();

            LogToFile("Scene loaded: " + sceneName + " (" + buildIndex.ToString(CultureInfo.InvariantCulture) + ")");
        }

        public override void OnUpdate()
        {
            try
            {
                if (Time.time < _nextScanTime)
                    return;

                _nextScanTime = Time.time + ScanIntervalSeconds;
                RunSaveDataBundlePass();
            }
            catch (Exception ex)
            {
                MelonLogger.Error("[AutoSwitch] OnUpdate exception: " + ex);
                LogToFile("OnUpdate exception: " + ex);
                _nextScanTime = Time.time + 10.0f;
            }
        }

        private void RunSaveDataBundlePass()
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

            NetworkSaveData networkSaveData = GetNetworkSaveData();
            List<CableSaveInfo> allSaveCables = ReadAllSaveCables(networkSaveData);

            FabricIdToSpeed.Clear();
            FabricIdToBundleCableIds.Clear();

            List<BundleBuilder> allBundles = new List<BundleBuilder>();
            HashSet<string> allManagedIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            int fabricIndex = 1;

            foreach (List<RegisteredSwitchInfo> fabric in fabrics)
            {
                string fabricId = "FABRIC-" + fabricIndex.ToString("000", CultureInfo.InvariantCulture);
                float fabricSpeed = fabric.Sum(x => x.EstimatedCapacityGbps);
                FabricIdToSpeed[fabricId] = fabricSpeed;

                HashSet<string> localIds = new HashSet<string>(
                    fabric.Select(x => x.DeviceName).Where(x => !string.IsNullOrWhiteSpace(x)),
                    StringComparer.OrdinalIgnoreCase
                );

                foreach (string localId in localIds)
                    allManagedIds.Add(localId);

                Dictionary<string, BundleBuilder> bundlesForFabric = BuildBundlesForFabric(localIds, allSaveCables);
                List<int> fabricCableIds = new List<int>();

                foreach (BundleBuilder bundle in bundlesForFabric.Values.OrderBy(b => b.LocalDeviceId).ThenBy(b => b.RemoteDeviceId))
                {
                    bundle.FabricId = fabricId;
                    bundle.FabricEstimatedSpeedGbps = fabricSpeed;

                    List<int> distinctIds = bundle.CableIds.Distinct().OrderBy(x => x).ToList();
                    if (distinctIds.Count >= 2)
                    {
                        foreach (int cableId in distinctIds)
                            fabricCableIds.Add(cableId);

                        allBundles.Add(bundle);

                        string signature =
                            fabricId + "|" +
                            bundle.LocalDeviceId + "|" +
                            bundle.RemoteDeviceId + "|" +
                            string.Join(",", distinctIds);

                        if (LoggedBundleSignatures.Add(signature))
                        {
                            LogToFile("SAVE BUNDLE | fabricId=" + fabricId +
                                      " | local=" + bundle.LocalDeviceId +
                                      " | remote=" + bundle.RemoteDeviceId +
                                      " | cableCount=" + distinctIds.Count.ToString(CultureInfo.InvariantCulture) +
                                      " | estPerCableGbps=" + (fabricSpeed / distinctIds.Count).ToString("0.##", CultureInfo.InvariantCulture) +
                                      " | cableIds=[" + string.Join(",", distinctIds) + "]");
                        }
                    }
                }

                FabricIdToBundleCableIds[fabricId] = fabricCableIds;
                fabricIndex++;
            }

            SaveDataAutoLACP.UpdateDesiredState(allBundles, allManagedIds);

            int adjacencyPairs = fabrics.Sum(f => Math.Max(0, f.Count - 1));
            string summary =
                "SCAN SUMMARY | registeredSwitches=" + RegisteredSwitches.Count.ToString(CultureInfo.InvariantCulture) +
                " | registeredCables=" + RegisteredCables.Count.ToString(CultureInfo.InvariantCulture) +
                " | saveCables=" + allSaveCables.Count.ToString(CultureInfo.InvariantCulture) +
                " | liveSwitches=" + liveSwitches.Count.ToString(CultureInfo.InvariantCulture) +
                " | activeFabrics=" + fabrics.Count.ToString(CultureInfo.InvariantCulture) +
                " | adjacentPairs=" + adjacencyPairs.ToString(CultureInfo.InvariantCulture) +
                " | bundleCount=" + allBundles.Count.ToString(CultureInfo.InvariantCulture) +
                " | managedIds=" + allManagedIds.Count.ToString(CultureInfo.InvariantCulture);

            if (!string.Equals(summary, _lastSummary, StringComparison.Ordinal))
            {
                _lastSummary = summary;
                MelonLogger.Msg("[AutoSwitch] " + summary);
                LogToFile(summary);

                int logIndex = 1;
                foreach (List<RegisteredSwitchInfo> fabric in fabrics)
                {
                    string fabricId = "FABRIC-" + logIndex.ToString("000", CultureInfo.InvariantCulture);
                    float speed = FabricIdToSpeed.ContainsKey(fabricId)
                        ? FabricIdToSpeed[fabricId]
                        : fabric.Sum(x => x.EstimatedCapacityGbps);

                    List<int> cableIds;
                    int tracked = FabricIdToBundleCableIds.TryGetValue(fabricId, out cableIds)
                        ? cableIds.Distinct().Count()
                        : 0;

                    string members = string.Join(", ",
                        fabric.Select(x => x.DeviceName + "@(" +
                                           x.WorldPosition.x.ToString("0.###", CultureInfo.InvariantCulture) + "," +
                                           x.WorldPosition.y.ToString("0.###", CultureInfo.InvariantCulture) + "," +
                                           x.WorldPosition.z.ToString("0.###", CultureInfo.InvariantCulture) + ")[" +
                                           x.ModelName + "]"));

                    LogToFile("FABRIC | id=" + fabricId +
                              " | members=" + fabric.Count.ToString(CultureInfo.InvariantCulture) +
                              " | estCapacityGbps=" + speed.ToString("0.##", CultureInfo.InvariantCulture) +
                              " | bundledCableIds=" + tracked.ToString(CultureInfo.InvariantCulture) +
                              " | switches=" + members);

                    if (tracked > 0)
                        LogToFile("FABRIC BUNDLE CABLE MAP | fabricId=" + fabricId + " | cableIds=[" + string.Join(",", cableIds.Distinct().OrderBy(x => x)) + "]");

                    logIndex++;
                }
            }
        }

        private void InstallNativePatches()
        {
            try
            {
                Type networkMapType = AccessTools.TypeByName("Il2Cpp.NetworkMap");
                if (networkMapType != null)
                {
                    PatchIfFound(networkMapType, "RegisterSwitch", nameof(NetworkMap_RegisterSwitch_Postfix));
                    PatchIfFound(networkMapType, "RegisterCableConnection", nameof(NetworkMap_RegisterCableConnection_Postfix));
                }
            }
            catch (Exception ex)
            {
                LogToFile("InstallNativePatches failed: " + ex);
            }
        }

        private void PatchIfFound(Type type, string methodName, string postfixName)
        {
            try
            {
                MethodInfo target = AccessTools.Method(type, methodName);
                if (target == null)
                {
                    LogToFile("PATCH MISS | " + type.FullName + "." + methodName);
                    return;
                }

                MethodInfo postfixMethod = typeof(AutoSwitchMod).GetMethod(postfixName, BindingFlags.NonPublic | BindingFlags.Static);
                if (postfixMethod == null)
                {
                    LogToFile("PATCH MISS POSTFIX | " + postfixName);
                    return;
                }

                HarmonyInstance.Patch(target, postfix: new HarmonyMethod(postfixMethod));
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
            }
            catch (Exception ex)
            {
                LogToFile("NetworkMap_RegisterCableConnection_Postfix failed: " + ex);
            }
        }

        internal static NetworkSaveData GetNetworkSaveData()
        {
            try
            {
                MainGameManager mgm = MainGameManager.instance;
                if (mgm == null)
                    return null;

                object[] roots =
                {
                    TryGetMemberValue(mgm, "currentSaveData"),
                    TryGetMemberValue(mgm, "saveData"),
                    TryGetMemberValue(mgm, "loadedSaveData"),
                    TryGetMemberValue(mgm, "currentSave"),
                };

                foreach (object root in roots)
                {
                    NetworkSaveData data = ExtractNetworkSaveData(root);
                    if (data != null)
                        return data;
                }
            }
            catch (Exception ex)
            {
                LogToFile("GetNetworkSaveData failed: " + ex);
            }

            return null;
        }

        private static NetworkSaveData ExtractNetworkSaveData(object root)
        {
            if (root == null)
                return null;

            NetworkSaveData direct = root as NetworkSaveData;
            if (direct != null)
                return direct;

            object[] nested =
            {
                TryGetMemberValue(root, "networkSaveData"),
                TryGetMemberValue(root, "networkData"),
                TryGetMemberValue(root, "networkSave"),
            };

            foreach (object candidate in nested)
            {
                NetworkSaveData cast = candidate as NetworkSaveData;
                if (cast != null)
                    return cast;
            }

            return null;
        }

        private static List<CableSaveInfo> ReadAllSaveCables(NetworkSaveData networkSaveData)
        {
            List<CableSaveInfo> result = new List<CableSaveInfo>();
            if (networkSaveData == null || networkSaveData.cables == null)
                return result;

            foreach (CableSaveData cable in networkSaveData.cables)
            {
                if (cable == null)
                    continue;

                result.Add(new CableSaveInfo
                {
                    CableId = cable.cableID,
                    StartToken = EndpointToken(cable.startPoint),
                    EndToken = EndpointToken(cable.endPoint)
                });
            }

            return result;
        }

        private static Dictionary<string, BundleBuilder> BuildBundlesForFabric(
            HashSet<string> localIds,
            List<CableSaveInfo> allCables)
        {
            Dictionary<string, BundleBuilder> result =
                new Dictionary<string, BundleBuilder>(StringComparer.OrdinalIgnoreCase);

            foreach (CableSaveInfo cable in allCables)
            {
                if (cable == null)
                    continue;

                string localDeviceId;
                string remoteDeviceId;
                if (!TryResolveRemoteDevice(cable, localIds, out localDeviceId, out remoteDeviceId))
                    continue;

                if (string.IsNullOrWhiteSpace(localDeviceId) || string.IsNullOrWhiteSpace(remoteDeviceId))
                    continue;

                string key = localDeviceId + "||" + remoteDeviceId;
                BundleBuilder builder;
                if (!result.TryGetValue(key, out builder))
                {
                    builder = new BundleBuilder();
                    builder.LocalDeviceId = localDeviceId;
                    builder.RemoteDeviceId = remoteDeviceId;
                    result[key] = builder;
                }

                builder.CableIds.Add(cable.CableId);
            }

            return result;
        }

        private static bool TryResolveRemoteDevice(
            CableSaveInfo cable,
            HashSet<string> localIds,
            out string localDeviceId,
            out string remoteDeviceId)
        {
            localDeviceId = null;
            remoteDeviceId = null;

            if (cable == null)
                return false;

            bool startIsLocal = !string.IsNullOrWhiteSpace(cable.StartToken) && localIds.Contains(cable.StartToken);
            bool endIsLocal = !string.IsNullOrWhiteSpace(cable.EndToken) && localIds.Contains(cable.EndToken);

            if (startIsLocal && !string.IsNullOrWhiteSpace(cable.EndToken))
            {
                localDeviceId = cable.StartToken;
                remoteDeviceId = cable.EndToken;
                return true;
            }

            if (endIsLocal && !string.IsNullOrWhiteSpace(cable.StartToken))
            {
                localDeviceId = cable.EndToken;
                remoteDeviceId = cable.StartToken;
                return true;
            }

            return false;
        }

        internal static bool GroupTouchesManagedIds(LACPGroupSaveData group, HashSet<string> managedIds)
        {
            if (group == null)
                return false;

            return (!string.IsNullOrWhiteSpace(group.deviceA) && managedIds.Contains(group.deviceA)) ||
                   (!string.IsNullOrWhiteSpace(group.deviceB) && managedIds.Contains(group.deviceB));
        }

        internal static string EndpointToken(CableEndpointSaveData endpoint)
        {
            if (endpoint == null)
                return null;

            if (!string.IsNullOrWhiteSpace(endpoint.switchID))
                return endpoint.switchID;

            if (!string.IsNullOrWhiteSpace(endpoint.serverID))
                return endpoint.serverID;

            if (endpoint.customerID >= 0)
                return "Customer_" + endpoint.customerID.ToString(CultureInfo.InvariantCulture);

            return null;
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

            return AllowedSwitchModels.Contains(NormalizeCloneName(modelName));
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
            return Mathf.Abs(actualStep - FallbackExpectedStep) <= AdjacentYTolerance;
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
                return "SERVERFAMILY:" + NormalizeRuntimeIdentity(raw);

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
            catch { }

            try
            {
                Component comp = obj as Component;
                if (comp != null)
                    return comp.gameObject;
            }
            catch { }

            try
            {
                object go = TryGetMemberValue(obj, "gameObject");
                GameObject gameObject = go as GameObject;
                if (gameObject != null)
                    return gameObject;
            }
            catch { }

            try
            {
                object transform = TryGetMemberValue(obj, "transform");
                Transform t = transform as Transform;
                if (t != null)
                    return t.gameObject;
            }
            catch { }

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
                catch { }
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
                catch { }
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
            catch { }

            return null;
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

        internal static void LogToFile(string message)
        {
            try
            {
                Directory.CreateDirectory(DebugFolderPath);
                File.AppendAllText(DebugLogPath, "[" + DateTime.Now.ToString("HH:mm:ss.fff") + "] " + message + Environment.NewLine);
            }
            catch { }
        }
    }

    internal static class SaveDataAutoLACP
    {
        private static bool _bootstrapStarted;
        private static bool _regroupQueued;
        private static float _lastRunRealtime;
        private static string _lastAppliedSignature = string.Empty;

        private const float InitialDelaySeconds = 8.0f;
        private const float MinDelaySeconds = 0.35f;
        private const float MinGapBetweenRunsSeconds = 0.50f;

        private static readonly object _stateLock = new object();
        private static List<BundleBuilder> _desiredBundles = new List<BundleBuilder>();
        private static HashSet<string> _desiredManagedIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        internal static void StartSafeBootstrap()
        {
            if (_bootstrapStarted)
                return;

            _bootstrapStarted = true;
            MelonLogger.Msg("[AutoSwitch] Save-data Auto-LACP bootstrap armed.");
            MelonCoroutines.Start(BootstrapRoutine());
        }

        internal static void ResetForScene()
        {
            _bootstrapStarted = false;
            _regroupQueued = false;
            _lastRunRealtime = 0f;
            _lastAppliedSignature = string.Empty;

            lock (_stateLock)
            {
                _desiredBundles = new List<BundleBuilder>();
                _desiredManagedIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            }
        }

        private static IEnumerator BootstrapRoutine()
        {
            yield return new WaitForSeconds(InitialDelaySeconds);
            QueueRegroup("safe bootstrap");
        }

        internal static void UpdateDesiredState(List<BundleBuilder> bundles, HashSet<string> managedIds)
        {
            if (bundles == null)
                bundles = new List<BundleBuilder>();

            if (managedIds == null)
                managedIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            string signature = BuildDesiredSignature(bundles, managedIds);

            lock (_stateLock)
            {
                _desiredBundles = bundles
                    .Select(CloneBundle)
                    .OrderBy(b => b.FabricId)
                    .ThenBy(b => b.LocalDeviceId)
                    .ThenBy(b => b.RemoteDeviceId)
                    .ToList();

                _desiredManagedIds = new HashSet<string>(managedIds, StringComparer.OrdinalIgnoreCase);
            }

            if (!string.Equals(signature, _lastAppliedSignature, StringComparison.Ordinal))
            {
                _lastAppliedSignature = signature;
                QueueRegroup("desired save-data bundles changed");
            }
        }

        internal static void QueueRegroup(string reason)
        {
            if (_regroupQueued)
                return;

            _regroupQueued = true;
            MelonLogger.Msg("[AutoSwitch] Save-data Auto-LACP queued: " + reason);
            MelonCoroutines.Start(RunQueuedRegroup());
        }

        private static IEnumerator RunQueuedRegroup()
        {
            yield return new WaitForSeconds(MinDelaySeconds);

            float now = Time.realtimeSinceStartup;
            if (now - _lastRunRealtime < MinGapBetweenRunsSeconds)
            {
                _regroupQueued = false;
                yield return new WaitForSeconds(MinGapBetweenRunsSeconds);
                QueueRegroup("debounced");
                yield break;
            }

            _regroupQueued = false;
            _lastRunRealtime = Time.realtimeSinceStartup;

            try
            {
                RebuildAutoLacpGroups();
            }
            catch (Exception ex)
            {
                MelonLogger.Warning("[AutoSwitch] Save-data Auto-LACP regroup failed: " + ex);
                AutoSwitchMod.LogToFile("Save-data Auto-LACP regroup failed: " + ex);
            }
        }

        internal static void RebuildAutoLacpGroups()
        {
            List<BundleBuilder> desiredBundles;
            HashSet<string> desiredManagedIds;

            lock (_stateLock)
            {
                desiredBundles = _desiredBundles.Select(CloneBundle).ToList();
                desiredManagedIds = new HashSet<string>(_desiredManagedIds, StringComparer.OrdinalIgnoreCase);
            }

            NetworkMap networkMap = NetworkMap.instance;
            NetworkSaveData networkSaveData = AutoSwitchMod.GetNetworkSaveData();

            if (networkMap == null)
            {
                AutoSwitchMod.LogToFile("AUTO LACP | skipped, NetworkMap.instance is null.");
                return;
            }

            List<int> removedGroupIds = new List<int>();

            if (networkSaveData != null && networkSaveData.lacpGroups != null)
            {
                foreach (LACPGroupSaveData existing in networkSaveData.lacpGroups)
                {
                    if (AutoSwitchMod.GroupTouchesManagedIds(existing, desiredManagedIds))
                        removedGroupIds.Add(existing.groupId);
                }
            }

            foreach (int groupId in removedGroupIds.Distinct().OrderBy(x => x))
            {
                try
                {
                    networkMap.RemoveLACPGroup(groupId);
                }
                catch (Exception ex)
                {
                    AutoSwitchMod.LogToFile("AUTO LACP | RemoveLACPGroup failed for " + groupId.ToString(CultureInfo.InvariantCulture) + " | " + ex.Message);
                }
            }

            if (networkSaveData != null)
            {
                List<LACPGroupSaveData> rebuiltSaveGroups = new List<LACPGroupSaveData>();

                if (networkSaveData.lacpGroups != null)
                {
                    foreach (LACPGroupSaveData existing in networkSaveData.lacpGroups)
                    {
                        if (!AutoSwitchMod.GroupTouchesManagedIds(existing, desiredManagedIds))
                            rebuiltSaveGroups.Add(existing);
                    }
                }

                foreach (BundleBuilder bundle in desiredBundles)
                {
                    List<int> distinctIds = bundle.CableIds.Distinct().OrderBy(x => x).ToList();
                    if (distinctIds.Count < 2)
                        continue;

                    try
                    {
                        Il2CppSystem.Collections.Generic.List<int> il2cppCableIds = new Il2CppSystem.Collections.Generic.List<int>();
                        foreach (int cableId in distinctIds)
                            il2cppCableIds.Add(cableId);

                        int groupId = networkMap.CreateLACPGroup(bundle.LocalDeviceId, bundle.RemoteDeviceId, il2cppCableIds);

                        try
                        {
                            LACPGroupSaveData saveGroup = new LACPGroupSaveData();
                            saveGroup.groupId = groupId;
                            saveGroup.deviceA = bundle.LocalDeviceId;
                            saveGroup.deviceB = bundle.RemoteDeviceId;
                            saveGroup.cableIds = il2cppCableIds;
                            rebuiltSaveGroups.Add(saveGroup);
                        }
                        catch (Exception ex)
                        {
                            AutoSwitchMod.LogToFile("AUTO LACP | save group construct failed for " +
                                                    bundle.LocalDeviceId + " -> " + bundle.RemoteDeviceId + " | " + ex.Message);
                        }

                        AutoSwitchMod.LogToFile("AUTO LACP | created groupId=" + groupId.ToString(CultureInfo.InvariantCulture) +
                                                " | fabricId=" + bundle.FabricId +
                                                " | deviceA=" + bundle.LocalDeviceId +
                                                " | deviceB=" + bundle.RemoteDeviceId +
                                                " | cableIds=[" + string.Join(",", distinctIds) + "]");
                    }
                    catch (Exception ex)
                    {
                        AutoSwitchMod.LogToFile("AUTO LACP | CreateLACPGroup failed | deviceA=" + bundle.LocalDeviceId +
                                                " | deviceB=" + bundle.RemoteDeviceId +
                                                " | " + ex);
                    }
                }

                try
                {
                    Il2CppSystem.Collections.Generic.List<LACPGroupSaveData> saveList = new Il2CppSystem.Collections.Generic.List<LACPGroupSaveData>();
                    foreach (LACPGroupSaveData group in rebuiltSaveGroups)
                        saveList.Add(group);

                    networkSaveData.lacpGroups = saveList;
                }
                catch (Exception ex)
                {
                    AutoSwitchMod.LogToFile("AUTO LACP | failed updating networkSaveData.lacpGroups | " + ex);
                }
            }

            AutoSwitchMod.LogToFile("AUTO LACP | regroup complete | removed=" +
                                    removedGroupIds.Distinct().Count().ToString(CultureInfo.InvariantCulture) +
                                    " | desiredBundles=" + desiredBundles.Count.ToString(CultureInfo.InvariantCulture));
        }

        private static string BuildDesiredSignature(List<BundleBuilder> bundles, HashSet<string> managedIds)
        {
            List<string> parts = new List<string>();

            foreach (BundleBuilder bundle in bundles
                .OrderBy(b => b.FabricId)
                .ThenBy(b => b.LocalDeviceId)
                .ThenBy(b => b.RemoteDeviceId))
            {
                parts.Add(bundle.FabricId + "|" +
                          bundle.LocalDeviceId + "|" +
                          bundle.RemoteDeviceId + "|" +
                          string.Join(",", bundle.CableIds.Distinct().OrderBy(x => x)));
            }

            parts.Add("managed:" + string.Join(",", managedIds.OrderBy(x => x)));

            return string.Join(" || ", parts);
        }

        private static BundleBuilder CloneBundle(BundleBuilder source)
        {
            BundleBuilder clone = new BundleBuilder();
            clone.FabricId = source.FabricId;
            clone.FabricEstimatedSpeedGbps = source.FabricEstimatedSpeedGbps;
            clone.LocalDeviceId = source.LocalDeviceId;
            clone.RemoteDeviceId = source.RemoteDeviceId;

            foreach (int cableId in source.CableIds)
                clone.CableIds.Add(cableId);

            return clone;
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

    internal sealed class CableSaveInfo
    {
        public int CableId;
        public string StartToken;
        public string EndToken;
    }

    internal sealed class BundleBuilder
    {
        public string FabricId;
        public float FabricEstimatedSpeedGbps;
        public string LocalDeviceId;
        public string RemoteDeviceId;
        public readonly List<int> CableIds = new List<int>();
    }
}