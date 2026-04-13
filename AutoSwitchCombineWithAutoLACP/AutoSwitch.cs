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

[assembly: MelonInfo(typeof(AutoSwitch.AutoSwitchMod), "AutoSwitch", "2.16.2", "Big Texas Jerky")]
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

        private static readonly HashSet<string> LoggedRemoteResolutionSignatures =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private static readonly HashSet<string> LoggedCrossFabricSignatures =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private static readonly HashSet<string> LoggedUnknownRemotePassThroughSignatures =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private float _nextScanTime;
        private string _lastSummary = string.Empty;

        public override void OnInitializeMelon()
        {
            ClassInjector.RegisterTypeInIl2Cpp<FabricGroupTag>();

            Directory.CreateDirectory(DebugFolderPath);
            File.WriteAllText(
                DebugLogPath,
                "[" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + "] AutoSwitch 2.16.2 debug log started." + Environment.NewLine
            );

            InstallNativePatches();

            MelonLogger.Msg("[AutoSwitch] v2.16.2 active. Unknown remote switches pass-through mode.");
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
            LoggedRemoteResolutionSignatures.Clear();
            LoggedCrossFabricSignatures.Clear();
            LoggedUnknownRemotePassThroughSignatures.Clear();

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
                RunLiveBundlePass();
            }
            catch (Exception ex)
            {
                MelonLogger.Error("[AutoSwitch] OnUpdate exception: " + ex);
                LogToFile("OnUpdate exception: " + ex);
                _nextScanTime = Time.time + 10.0f;
            }
        }

        private void RunLiveBundlePass()
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

            Dictionary<string, string> deviceIdToFabricId = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < fabrics.Count; i++)
            {
                string fabricId = "FABRIC-" + (i + 1).ToString("000", CultureInfo.InvariantCulture);
                foreach (RegisteredSwitchInfo sw in fabrics[i])
                {
                    if (sw != null && !string.IsNullOrWhiteSpace(sw.DeviceName))
                        deviceIdToFabricId[sw.DeviceName] = fabricId;
                }
            }

            NetworkSaveData networkSaveData = GetNetworkSaveData();
            int saveCableCount = CountSaveCables(networkSaveData);

            HashSet<int> externalLacpCableIds = GetExternalLacpCableIds(networkSaveData);

            FabricIdToSpeed.Clear();
            FabricIdToBundleCableIds.Clear();

            List<BundleBuilder> allBundles = new List<BundleBuilder>();
            HashSet<string> allManagedIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            HashSet<int> globallyClaimedCableIds = new HashSet<int>();

            int crossFabricSkipCount = 0;
            HashSet<int> crossFabricSkippedCableIds = new HashSet<int>();

            int unknownRemotePassThroughCount = 0;
            HashSet<int> unknownRemotePassThroughCableIds = new HashSet<int>();

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

                Dictionary<string, BundleBuilder> bundlesForFabric =
                    BuildBundlesForFabricFromLiveRegistry(
                        fabricId,
                        localIds,
                        deviceIdToFabricId,
                        globallyClaimedCableIds,
                        externalLacpCableIds,
                        ref crossFabricSkipCount,
                        crossFabricSkippedCableIds,
                        ref unknownRemotePassThroughCount,
                        unknownRemotePassThroughCableIds);

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
                            LogToFile(
                                "LIVE BUNDLE | fabricId=" + fabricId +
                                " | local=" + bundle.LocalDeviceId +
                                " | remote=" + bundle.RemoteDeviceId +
                                " | cableCount=" + distinctIds.Count.ToString(CultureInfo.InvariantCulture) +
                                " | estPerCableGbps=" + (fabricSpeed / distinctIds.Count).ToString("0.##", CultureInfo.InvariantCulture) +
                                " | cableIds=[" + string.Join(",", distinctIds) + "]"
                            );
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
                " | saveCables=" + saveCableCount.ToString(CultureInfo.InvariantCulture) +
                " | externalLacpCableIds=" + externalLacpCableIds.Count.ToString(CultureInfo.InvariantCulture) +
                " | crossFabricSkipped=" + crossFabricSkipCount.ToString(CultureInfo.InvariantCulture) +
                " | unknownRemotePassThrough=" + unknownRemotePassThroughCount.ToString(CultureInfo.InvariantCulture) +
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

                if (externalLacpCableIds.Count > 0)
                    LogToFile("EXTERNAL LACP CABLE IDS | [" + string.Join(",", externalLacpCableIds.OrderBy(x => x)) + "]");

                if (crossFabricSkippedCableIds.Count > 0)
                    LogToFile("CROSS FABRIC SKIPPED CABLE IDS | [" + string.Join(",", crossFabricSkippedCableIds.OrderBy(x => x)) + "]");

                if (unknownRemotePassThroughCableIds.Count > 0)
                    LogToFile("UNKNOWN REMOTE PASS-THROUGH CABLE IDS | [" + string.Join(",", unknownRemotePassThroughCableIds.OrderBy(x => x)) + "]");

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

                    LogToFile(
                        "FABRIC | id=" + fabricId +
                        " | members=" + fabric.Count.ToString(CultureInfo.InvariantCulture) +
                        " | estCapacityGbps=" + speed.ToString("0.##", CultureInfo.InvariantCulture) +
                        " | bundledCableIds=" + tracked.ToString(CultureInfo.InvariantCulture) +
                        " | switches=" + members
                    );

                    if (tracked > 0)
                    {
                        LogToFile(
                            "FABRIC BUNDLE CABLE MAP | fabricId=" + fabricId +
                            " | cableIds=[" + string.Join(",", cableIds.Distinct().OrderBy(x => x)) + "]"
                        );
                    }

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
                info.EndpointSwitchIdA = ResolveEndpointSwitchId(__3, __5, __9);
                info.EndpointSwitchIdB = ResolveEndpointSwitchId(__4, __6, __10);

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

        private static int CountSaveCables(NetworkSaveData networkSaveData)
        {
            try
            {
                if (networkSaveData == null || networkSaveData.cables == null)
                    return 0;

                int count = 0;
                foreach (CableSaveData cable in networkSaveData.cables)
                {
                    if (cable != null)
                        count++;
                }

                return count;
            }
            catch (Exception ex)
            {
                LogToFile("CountSaveCables failed: " + ex);
                return 0;
            }
        }

        private static HashSet<int> GetExternalLacpCableIds(NetworkSaveData networkSaveData)
        {
            HashSet<int> result = new HashSet<int>();

            try
            {
                if (networkSaveData == null || networkSaveData.lacpGroups == null)
                    return result;

                HashSet<int> owned = SaveDataAutoLACP.GetOwnedGroupIdsSnapshot();

                foreach (LACPGroupSaveData group in networkSaveData.lacpGroups)
                {
                    if (group == null)
                        continue;

                    if (owned.Contains(group.groupId))
                        continue;

                    if (group.cableIds == null)
                        continue;

                    foreach (int cableId in group.cableIds)
                        result.Add(cableId);
                }
            }
            catch (Exception ex)
            {
                LogToFile("GetExternalLacpCableIds failed: " + ex);
            }

            return result;
        }

        private static Dictionary<string, BundleBuilder> BuildBundlesForFabricFromLiveRegistry(
            string currentFabricId,
            HashSet<string> localIds,
            Dictionary<string, string> deviceIdToFabricId,
            HashSet<int> globallyClaimedCableIds,
            HashSet<int> externalLacpCableIds,
            ref int crossFabricSkipCount,
            HashSet<int> crossFabricSkippedCableIds,
            ref int unknownRemotePassThroughCount,
            HashSet<int> unknownRemotePassThroughCableIds)
        {
            Dictionary<string, BundleBuilder> result =
                new Dictionary<string, BundleBuilder>(StringComparer.OrdinalIgnoreCase);

            foreach (RegisteredCableInfo cable in RegisteredCables.Values.OrderBy(x => x.CableId))
            {
                if (cable == null)
                    continue;

                if (globallyClaimedCableIds.Contains(cable.CableId))
                    continue;

                if (externalLacpCableIds.Contains(cable.CableId))
                    continue;

                string localDeviceId;
                string remoteDeviceId;
                if (!TryResolveRemoteDeviceFromLiveCable(cable, localIds, out localDeviceId, out remoteDeviceId))
                    continue;

                if (string.IsNullOrWhiteSpace(localDeviceId) || string.IsNullOrWhiteSpace(remoteDeviceId))
                    continue;

                bool remoteIsSwitch = LooksLikeSwitchIdentity(remoteDeviceId);
                if (remoteIsSwitch)
                {
                    string remoteFabricId;
                    if (deviceIdToFabricId.TryGetValue(remoteDeviceId, out remoteFabricId))
                    {
                        if (!string.Equals(remoteFabricId, currentFabricId, StringComparison.OrdinalIgnoreCase))
                        {
                            crossFabricSkipCount++;
                            crossFabricSkippedCableIds.Add(cable.CableId);
                            LogCrossFabricSkip(cable.CableId, currentFabricId, localDeviceId, remoteDeviceId, remoteFabricId);
                            continue;
                        }
                    }
                    else
                    {
                        unknownRemotePassThroughCount++;
                        unknownRemotePassThroughCableIds.Add(cable.CableId);
                        LogUnknownRemotePassThrough(cable.CableId, currentFabricId, localDeviceId, remoteDeviceId);
                        continue;
                    }
                }

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
                globallyClaimedCableIds.Add(cable.CableId);
            }

            return result;
        }

        private static void LogCrossFabricSkip(
            int cableId,
            string localFabricId,
            string localDeviceId,
            string remoteDeviceId,
            string remoteFabricId)
        {
            string signature =
                cableId.ToString(CultureInfo.InvariantCulture) + "|" +
                localFabricId + "|" +
                localDeviceId + "|" +
                remoteDeviceId + "|" +
                remoteFabricId;

            if (!LoggedCrossFabricSignatures.Add(signature))
                return;

            LogToFile(
                "CROSS FABRIC SKIP | cableId=" + cableId.ToString(CultureInfo.InvariantCulture) +
                " | localFabric=" + localFabricId +
                " | localDevice=" + localDeviceId +
                " | remoteDevice=" + remoteDeviceId +
                " | remoteFabric=" + remoteFabricId
            );
        }

        private static void LogUnknownRemotePassThrough(
            int cableId,
            string localFabricId,
            string localDeviceId,
            string remoteDeviceId)
        {
            string signature =
                cableId.ToString(CultureInfo.InvariantCulture) + "|" +
                localFabricId + "|" +
                localDeviceId + "|" +
                remoteDeviceId;

            if (!LoggedUnknownRemotePassThroughSignatures.Add(signature))
                return;

            LogToFile(
                "UNKNOWN REMOTE PASS-THROUGH | cableId=" + cableId.ToString(CultureInfo.InvariantCulture) +
                " | localFabric=" + localFabricId +
                " | localDevice=" + localDeviceId +
                " | remoteDevice=" + remoteDeviceId
            );
        }

        private static bool TryResolveRemoteDeviceFromLiveCable(
            RegisteredCableInfo cable,
            HashSet<string> localIds,
            out string localDeviceId,
            out string remoteDeviceId)
        {
            localDeviceId = null;
            remoteDeviceId = null;

            if (cable == null)
                return false;

            bool aIsLocal = !string.IsNullOrWhiteSpace(cable.DeviceA) && localIds.Contains(cable.DeviceA);
            bool bIsLocal = !string.IsNullOrWhiteSpace(cable.DeviceB) && localIds.Contains(cable.DeviceB);

            if (aIsLocal && bIsLocal)
                return false;

            if (aIsLocal)
            {
                localDeviceId = cable.DeviceA;
                remoteDeviceId = BuildRemoteIdentity(
                    cable.CableId,
                    localDeviceId,
                    cable.DeviceB,
                    cable.ExtraB,
                    cable.ServerRootKeyB,
                    cable.EndpointSwitchIdB,
                    cable.PortB);
                return !string.IsNullOrWhiteSpace(remoteDeviceId);
            }

            if (bIsLocal)
            {
                localDeviceId = cable.DeviceB;
                remoteDeviceId = BuildRemoteIdentity(
                    cable.CableId,
                    localDeviceId,
                    cable.DeviceA,
                    cable.ExtraA,
                    cable.ServerRootKeyA,
                    cable.EndpointSwitchIdA,
                    cable.PortA);
                return !string.IsNullOrWhiteSpace(remoteDeviceId);
            }

            return false;
        }

        private static string BuildRemoteIdentity(
            int cableId,
            string localDeviceId,
            string deviceName,
            string extraName,
            string serverRootKey,
            string endpointSwitchId,
            int port)
        {
            string preservedDevice = PreserveSwitchIdentity(deviceName);
            string preservedEndpointSwitch = PreserveSwitchIdentity(endpointSwitchId);
            string normalizedExtra = NormalizeRuntimeIdentity(extraName);

            if (!string.IsNullOrWhiteSpace(preservedDevice))
            {
                LogRemoteResolution(cableId, localDeviceId, "deviceId", preservedDevice, deviceName, extraName, endpointSwitchId, serverRootKey, port);
                return preservedDevice;
            }

            if (!string.IsNullOrWhiteSpace(preservedEndpointSwitch))
            {
                LogRemoteResolution(cableId, localDeviceId, "endpointSwitch", preservedEndpointSwitch, deviceName, extraName, endpointSwitchId, serverRootKey, port);
                return preservedEndpointSwitch;
            }

            if (!string.IsNullOrWhiteSpace(serverRootKey))
            {
                LogRemoteResolution(cableId, localDeviceId, "serverRoot", serverRootKey, deviceName, extraName, endpointSwitchId, serverRootKey, port);
                return serverRootKey;
            }

            if (!string.IsNullOrWhiteSpace(normalizedExtra))
            {
                string value = normalizedExtra + "#P" + port.ToString(CultureInfo.InvariantCulture);
                LogRemoteResolution(cableId, localDeviceId, "fallbackExtra", value, deviceName, extraName, endpointSwitchId, serverRootKey, port);
                return value;
            }

            return string.Empty;
        }

        private static string PreserveSwitchIdentity(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
                return string.Empty;

            string trimmed = raw.Trim();

            if (LooksLikeSwitchIdentity(trimmed))
                return trimmed;

            return string.Empty;
        }

        private static bool LooksLikeSwitchIdentity(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
                return false;

            string trimmed = raw.Trim();

            return trimmed.StartsWith("Switch32xQSFP", StringComparison.OrdinalIgnoreCase) ||
                   trimmed.StartsWith("Switch4xQSXP16xSFP", StringComparison.OrdinalIgnoreCase) ||
                   trimmed.StartsWith("Switch4xSFP", StringComparison.OrdinalIgnoreCase) ||
                   trimmed.StartsWith("Switch16CU", StringComparison.OrdinalIgnoreCase);
        }

        private static void LogRemoteResolution(
            int cableId,
            string localDeviceId,
            string kind,
            string chosen,
            string deviceName,
            string extraName,
            string endpointSwitchId,
            string serverRootKey,
            int port)
        {
            string signature =
                cableId.ToString(CultureInfo.InvariantCulture) + "|" +
                localDeviceId + "|" +
                kind + "|" +
                chosen;

            if (!LoggedRemoteResolutionSignatures.Add(signature))
                return;

            LogToFile(
                "REMOTE RESOLVE | cableId=" + cableId.ToString(CultureInfo.InvariantCulture) +
                " | local=" + localDeviceId +
                " | kind=" + kind +
                " | chosen=" + chosen +
                " | rawDevice=" + (deviceName ?? string.Empty) +
                " | rawExtra=" + (extraName ?? string.Empty) +
                " | endpointSwitchId=" + (endpointSwitchId ?? string.Empty) +
                " | serverRoot=" + (serverRootKey ?? string.Empty) +
                " | port=" + port.ToString(CultureInfo.InvariantCulture)
            );
        }

        private static string ResolveEndpointSwitchId(object endpointObj, string deviceName, string extraName)
        {
            try
            {
                object[] candidates =
                {
                    endpointObj,
                    TryGetMemberValue(endpointObj, "connectedSwitch"),
                    TryGetMemberValue(endpointObj, "networkSwitch"),
                    TryGetMemberValue(endpointObj, "switchRef"),
                    TryGetMemberValue(endpointObj, "parentSwitch"),
                    TryGetMemberValue(endpointObj, "switchData")
                };

                foreach (object candidate in candidates)
                {
                    string discovered = DiscoverBestDeviceName(candidate);
                    if (!string.IsNullOrWhiteSpace(discovered) && IsSwitchLikeIdentity(discovered))
                        return discovered;

                    GameObject go = DiscoverGameObject(candidate);
                    string fromGo = ResolveSwitchIdFromGameObject(go);
                    if (!string.IsNullOrWhiteSpace(fromGo))
                        return fromGo;
                }

                GameObject endpointGo = DiscoverGameObject(endpointObj);
                string fromEndpointGo = ResolveSwitchIdFromGameObject(endpointGo);
                if (!string.IsNullOrWhiteSpace(fromEndpointGo))
                    return fromEndpointGo;

                string raw = !string.IsNullOrWhiteSpace(deviceName) ? deviceName : extraName;
                if (IsSwitchLikeIdentity(raw))
                    return raw;
            }
            catch (Exception ex)
            {
                LogToFile("ResolveEndpointSwitchId failed: " + ex.Message);
            }

            return string.Empty;
        }

        private static string ResolveSwitchIdFromGameObject(GameObject go)
        {
            if (go == null)
                return string.Empty;

            try
            {
                Transform current = go.transform;
                int depth = 0;

                while (current != null && depth < 20)
                {
                    string discovered = DiscoverBestDeviceName(current.gameObject);
                    if (!string.IsNullOrWhiteSpace(discovered) && IsSwitchLikeIdentity(discovered))
                        return discovered;

                    string currentName = NormalizeCloneName(current.name);
                    if (AllowedSwitchModels.Contains(currentName))
                    {
                        string modelNamed = DiscoverBestDeviceName(current.gameObject);
                        if (!string.IsNullOrWhiteSpace(modelNamed) && IsSwitchLikeIdentity(modelNamed))
                            return modelNamed;
                    }

                    current = current.parent;
                    depth++;
                }
            }
            catch { }

            return string.Empty;
        }

        private static bool IsSwitchLikeIdentity(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return false;

            string normalized = NormalizeCloneName(value);

            return normalized.StartsWith("Switch32xQSFP", StringComparison.OrdinalIgnoreCase) ||
                   normalized.StartsWith("Switch4xQSXP16xSFP", StringComparison.OrdinalIgnoreCase) ||
                   normalized.StartsWith("Switch4xSFP", StringComparison.OrdinalIgnoreCase) ||
                   normalized.StartsWith("Switch16CU", StringComparison.OrdinalIgnoreCase);
        }

        internal static bool GroupTouchesManagedIds(LACPGroupSaveData group, HashSet<string> managedIds)
        {
            if (group == null)
                return false;

            return (!string.IsNullOrWhiteSpace(group.deviceA) && managedIds.Contains(group.deviceA)) ||
                   (!string.IsNullOrWhiteSpace(group.deviceB) && managedIds.Contains(group.deviceB));
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
            info.EstimatedCapacityGbps = EstimateCapacityFromModel(modelName);

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
            if (obj == null)
                return string.Empty;

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

            Type type;
            try
            {
                type = obj.GetType();
            }
            catch
            {
                return string.Empty;
            }

            foreach (FieldInfo f in SafeFields(type))
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

            foreach (PropertyInfo p in SafeProperties(type))
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
                File.AppendAllText(
                    DebugLogPath,
                    "[" + DateTime.Now.ToString("HH:mm:ss.fff", CultureInfo.InvariantCulture) + "] " + message + Environment.NewLine
                );
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
        private static bool _hasAppliedAtLeastOnce;

        private const float InitialDelaySeconds = 8.0f;
        private const float MinDelaySeconds = 0.35f;
        private const float MinGapBetweenRunsSeconds = 0.50f;

        private static readonly object _stateLock = new object();
        private static List<BundleBuilder> _desiredBundles = new List<BundleBuilder>();
        private static HashSet<string> _desiredManagedIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private static readonly HashSet<int> _ownedGroupIds = new HashSet<int>();

        internal static void StartSafeBootstrap()
        {
            if (_bootstrapStarted)
                return;

            _bootstrapStarted = true;
            MelonLogger.Msg("[AutoSwitch] Live Auto-LACP bootstrap armed.");
            MelonCoroutines.Start(BootstrapRoutine());
        }

        internal static void ResetForScene()
        {
            _bootstrapStarted = false;
            _regroupQueued = false;
            _lastRunRealtime = 0f;
            _lastAppliedSignature = string.Empty;
            _hasAppliedAtLeastOnce = false;

            lock (_stateLock)
            {
                _desiredBundles = new List<BundleBuilder>();
                _desiredManagedIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            }

            _ownedGroupIds.Clear();
        }

        internal static HashSet<int> GetOwnedGroupIdsSnapshot()
        {
            return new HashSet<int>(_ownedGroupIds);
        }

        private static IEnumerator BootstrapRoutine()
        {
            yield return new WaitForSeconds(InitialDelaySeconds);

            if (_hasAppliedAtLeastOnce)
            {
                AutoSwitchMod.LogToFile("AUTO LACP | bootstrap skipped because desired state was already applied.");
                yield break;
            }

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
                QueueRegroup("desired live bundles changed");
            }
        }

        internal static void QueueRegroup(string reason)
        {
            if (_regroupQueued)
                return;

            _regroupQueued = true;
            MelonLogger.Msg("[AutoSwitch] Live Auto-LACP queued: " + reason);
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
                _hasAppliedAtLeastOnce = true;
            }
            catch (Exception ex)
            {
                MelonLogger.Warning("[AutoSwitch] Live Auto-LACP regroup failed: " + ex);
                AutoSwitchMod.LogToFile("Live Auto-LACP regroup failed: " + ex);
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

            foreach (int ownedId in _ownedGroupIds.OrderBy(x => x).ToList())
            {
                try
                {
                    networkMap.RemoveLACPGroup(ownedId);
                    removedGroupIds.Add(ownedId);
                }
                catch (Exception ex)
                {
                    AutoSwitchMod.LogToFile("AUTO LACP | Remove owned group failed for " +
                                            ownedId.ToString(CultureInfo.InvariantCulture) + " | " + ex.Message);
                }
            }

            _ownedGroupIds.Clear();

            if (networkSaveData != null && networkSaveData.lacpGroups != null)
            {
                foreach (LACPGroupSaveData existing in networkSaveData.lacpGroups)
                {
                    if (AutoSwitchMod.GroupTouchesManagedIds(existing, desiredManagedIds))
                    {
                        try
                        {
                            networkMap.RemoveLACPGroup(existing.groupId);
                            removedGroupIds.Add(existing.groupId);
                        }
                        catch (Exception ex)
                        {
                            AutoSwitchMod.LogToFile("AUTO LACP | Remove save group failed for " +
                                                    existing.groupId.ToString(CultureInfo.InvariantCulture) + " | " + ex.Message);
                        }
                    }
                }
            }

            List<LACPGroupSaveData> rebuiltSaveGroups = new List<LACPGroupSaveData>();

            if (networkSaveData != null && networkSaveData.lacpGroups != null)
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
                    _ownedGroupIds.Add(groupId);

                    if (networkSaveData != null)
                    {
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

            if (networkSaveData != null)
            {
                try
                {
                    Il2CppSystem.Collections.Generic.List<LACPGroupSaveData> saveList =
                        new Il2CppSystem.Collections.Generic.List<LACPGroupSaveData>();

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
                                    " | desiredBundles=" + desiredBundles.Count.ToString(CultureInfo.InvariantCulture) +
                                    " | ownedNow=" + _ownedGroupIds.Count.ToString(CultureInfo.InvariantCulture));
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

        public string EndpointSwitchIdA;
        public string EndpointSwitchIdB;
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