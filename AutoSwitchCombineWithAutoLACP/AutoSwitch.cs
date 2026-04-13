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

[assembly: MelonInfo(typeof(AutoSwitch.AutoSwitchMod), "AutoSwitch", "2.2.0", "Big Texas Jerky")]
[assembly: MelonGame("Waseku", "Data Center")]

namespace AutoSwitch
{
    public sealed class AutoSwitchMod : MelonMod
    {
        private const float ScanIntervalSeconds = 6.0f;

        private const float SameFabricXZTolerance = 0.40f;
        private const float AdjacentYTolerance = 0.040f;
        private const float FallbackExpectedStep = 0.0444f;

        // Safe-mode junk filters
        private const float OriginXZRejectRadius = 1.25f;
        private const float MinUsefulWorldY = 1.0f;
        private const float MaxUsefulWorldY = 4.5f;

        private static readonly Regex TrailingRuntimeIdRegex =
            new(@"_[\-]?\d+$", RegexOptions.Compiled);

        private static string DebugFolderPath =>
            Path.Combine(MelonEnvironment.ModsDirectory, "AutoSwitch");

        private static string DebugLogPath =>
            Path.Combine(DebugFolderPath, "autoswitch-debug.log");

        private static readonly HashSet<string> AllowedSwitchModels = new(StringComparer.OrdinalIgnoreCase)
        {
            "Switch32xQSFP",
            "Switch4xQSXP16xSFP",
            "Switch4xSFP",
            "Switch16CU"
        };

        private static readonly Dictionary<string, RegisteredSwitchInfo> RegisteredSwitches =
            new(StringComparer.OrdinalIgnoreCase);

        private static readonly Dictionary<int, RegisteredCableInfo> RegisteredCables =
            new();

        private static readonly Dictionary<string, float> FabricIdToSpeed =
            new(StringComparer.OrdinalIgnoreCase);

        private static readonly Dictionary<string, HashSet<int>> FabricIdToCableIds =
            new(StringComparer.OrdinalIgnoreCase);

        private static readonly Dictionary<int, string> CableIdToFabricId =
            new();

        private static readonly Dictionary<int, float> CableIdToOverrideSpeed =
            new();

        private static readonly Dictionary<string, string> LastFabricSignatures =
            new(StringComparer.OrdinalIgnoreCase);

        private static readonly HashSet<string> LoggedPatches =
            new(StringComparer.OrdinalIgnoreCase);

        private float _nextScanTime;
        private string _lastSummary = string.Empty;

        public override void OnInitializeMelon()
        {
            ClassInjector.RegisterTypeInIl2Cpp<FabricGroupTag>();

            Directory.CreateDirectory(DebugFolderPath);
            File.WriteAllText(DebugLogPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] AutoSwitch 2.2 debug log started.{Environment.NewLine}");

            InstallNativePatches();

            MelonLogger.Msg("[AutoSwitch] v2.2.0 active. Safe mode with normalized server bundles.");
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
            FabricIdToCableIds.Clear();
            CableIdToFabricId.Clear();
            CableIdToOverrideSpeed.Clear();
            LastFabricSignatures.Clear();

            LogToFile($"Scene loaded: {sceneName} ({buildIndex})");
        }

        public override void OnUpdate()
        {
            try
            {
                if (Time.time < _nextScanTime)
                    return;

                _nextScanTime = Time.time + ScanIntervalSeconds;
                RunSafeModeFabricPass();
            }
            catch (Exception ex)
            {
                MelonLogger.Error($"[AutoSwitch] OnUpdate exception: {ex}");
                LogToFile($"OnUpdate exception: {ex}");
                _nextScanTime = Time.time + 10.0f;
            }
        }

        private void RunSafeModeFabricPass()
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

            ApplySafeModeFabrics(fabrics);

            int adjacencyPairs = fabrics.Sum(f => Math.Max(0, f.Count - 1));

            string summary =
                $"SCAN SUMMARY | registeredSwitches={RegisteredSwitches.Count} | registeredCables={RegisteredCables.Count} | liveSwitches={liveSwitches.Count} | activeFabrics={fabrics.Count} | adjacentPairs={adjacencyPairs} | trackedCableIds={CableIdToOverrideSpeed.Count}";

            if (!string.Equals(summary, _lastSummary, StringComparison.Ordinal))
            {
                _lastSummary = summary;
                MelonLogger.Msg($"[AutoSwitch] {summary}");
                LogToFile(summary);

                int fabricIndex = 1;
                foreach (List<RegisteredSwitchInfo> fabric in fabrics)
                {
                    string fabricId = $"FABRIC-{fabricIndex:000}";
                    float speed = FabricIdToSpeed.TryGetValue(fabricId, out float s) ? s : fabric.Sum(x => x.EstimatedCapacityGbps);
                    int tracked = FabricIdToCableIds.TryGetValue(fabricId, out HashSet<int> ids) ? ids.Count : 0;

                    string members = string.Join(", ",
                        fabric.Select(x => $"{x.DeviceName}@({x.WorldPosition.x:0.###},{x.WorldPosition.y:0.###},{x.WorldPosition.z:0.###})[{x.ModelName}]"));

                    LogToFile($"FABRIC | id={fabricId} | members={fabric.Count} | estCapacityGbps={speed.ToString("0.##", CultureInfo.InvariantCulture)} | trackedCableIds={tracked} | switches={members}");

                    if (tracked > 0)
                    {
                        string cableList = string.Join(",", ids.OrderBy(x => x));
                        LogToFile($"FABRIC CABLE MAP | fabricId={fabricId} | cableIds=[{cableList}]");
                    }

                    fabricIndex++;
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
                    MethodInfo registerSwitch = AccessTools.Method(networkMapType, "RegisterSwitch");
                    if (registerSwitch != null)
                    {
                        HarmonyInstance.Patch(
                            registerSwitch,
                            postfix: new HarmonyMethod(typeof(AutoSwitchMod).GetMethod(nameof(NetworkMap_RegisterSwitch_Postfix), BindingFlags.NonPublic | BindingFlags.Static))
                        );
                        LogPatch("Patched NetworkMap.RegisterSwitch");
                    }

                    MethodInfo registerCable = AccessTools.Method(networkMapType, "RegisterCableConnection");
                    if (registerCable != null)
                    {
                        HarmonyInstance.Patch(
                            registerCable,
                            postfix: new HarmonyMethod(typeof(AutoSwitchMod).GetMethod(nameof(NetworkMap_RegisterCableConnection_Postfix), BindingFlags.NonPublic | BindingFlags.Static))
                        );
                        LogPatch("Patched NetworkMap.RegisterCableConnection");
                    }

                    MethodInfo createLacp = AccessTools.Method(networkMapType, "CreateLACPGroup");
                    if (createLacp != null)
                    {
                        HarmonyInstance.Patch(
                            createLacp,
                            postfix: new HarmonyMethod(typeof(AutoSwitchMod).GetMethod(nameof(NetworkMap_CreateLACPGroup_Postfix), BindingFlags.NonPublic | BindingFlags.Static))
                        );
                        LogPatch("Patched NetworkMap.CreateLACPGroup");
                    }
                }

                Type lacpGroupType = AccessTools.TypeByName("Il2Cpp.NetworkMap+LACPGroup");
                if (lacpGroupType != null)
                {
                    MethodInfo aggSpeed = AccessTools.Method(lacpGroupType, "GetAggregatedSpeed");
                    if (aggSpeed != null)
                    {
                        HarmonyInstance.Patch(
                            aggSpeed,
                            postfix: new HarmonyMethod(typeof(AutoSwitchMod).GetMethod(nameof(LACPGroup_GetAggregatedSpeed_Postfix), BindingFlags.NonPublic | BindingFlags.Static))
                        );
                        LogPatch("Patched NetworkMap+LACPGroup.GetAggregatedSpeed");
                    }
                }

                Type cableLinkType = AccessTools.TypeByName("Il2Cpp.CableLink");
                if (cableLinkType != null)
                {
                    MethodInfo setConnectionSpeed = AccessTools.Method(cableLinkType, "SetConnectionSpeed", new[] { typeof(float) });
                    if (setConnectionSpeed != null)
                    {
                        HarmonyInstance.Patch(
                            setConnectionSpeed,
                            prefix: new HarmonyMethod(typeof(AutoSwitchMod).GetMethod(nameof(CableLink_SetConnectionSpeed_Prefix), BindingFlags.NonPublic | BindingFlags.Static))
                        );
                        LogPatch("Patched CableLink.SetConnectionSpeed");
                    }
                }
            }
            catch (Exception ex)
            {
                LogToFile($"InstallNativePatches failed: {ex}");
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

                LogToFile(
                    $"REGISTER SWITCH | deviceName={info.DeviceName} | model={info.ModelName} | pos=({info.WorldPosition.x:0.###},{info.WorldPosition.y:0.###},{info.WorldPosition.z:0.###})"
                );
            }
            catch (Exception ex)
            {
                LogToFile($"NetworkMap_RegisterSwitch_Postfix failed: {ex}");
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
                var info = new RegisteredCableInfo
                {
                    CableId = __0,
                    DeviceA = (__5 ?? string.Empty).Trim(),
                    DeviceB = (__6 ?? string.Empty).Trim(),
                    ExtraA = (__9 ?? string.Empty).Trim(),
                    ExtraB = (__10 ?? string.Empty).Trim(),
                    PortA = __7,
                    PortB = __8
                };

                RegisteredCables[__0] = info;

                LogToFile(
                    $"REGISTER CABLE | cableId={info.CableId} | deviceA={info.DeviceA} | deviceB={info.DeviceB} | extraA={info.ExtraA} | extraB={info.ExtraB} | portA={info.PortA} | portB={info.PortB}"
                );
            }
            catch (Exception ex)
            {
                LogToFile($"NetworkMap_RegisterCableConnection_Postfix failed: {ex}");
            }
        }

        private static void NetworkMap_CreateLACPGroup_Postfix(string __0, string __1, object __2, ref int __result)
        {
            try
            {
                List<int> cableIds = ExtractIntList(__2);
                string joined = string.Join(",", cableIds.OrderBy(x => x));

                float sumSpeed = 0f;
                foreach (int cableId in cableIds)
                {
                    if (CableIdToOverrideSpeed.TryGetValue(cableId, out float speed))
                        sumSpeed += speed;
                }

                LogToFile($"CREATE LACP | groupId={__result} | deviceA={__0} | deviceB={__1} | cableIds=[{joined}] | overrideSum={sumSpeed.ToString("0.##", CultureInfo.InvariantCulture)}");
            }
            catch (Exception ex)
            {
                LogToFile($"NetworkMap_CreateLACPGroup_Postfix failed: {ex}");
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

                float overrideSum = 0f;
                foreach (int cableId in cableIds)
                {
                    if (CableIdToOverrideSpeed.TryGetValue(cableId, out float speed))
                        overrideSum += speed;
                }

                if (overrideSum > __result)
                    __result = overrideSum;
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

                object raw = TryGetMemberValue(__instance, "cableIDsOnLink");
                List<int> cableIds = ExtractIntList(raw);

                if (cableIds.Count == 0)
                    return;

                float overrideSpeed = 0f;
                foreach (int cableId in cableIds)
                {
                    if (CableIdToOverrideSpeed.TryGetValue(cableId, out float speed))
                        overrideSpeed = Mathf.Max(overrideSpeed, speed);
                }

                if (overrideSpeed > __0)
                    __0 = overrideSpeed;
            }
            catch
            {
            }
        }

        private static RegisteredSwitchInfo BuildRegisteredSwitchInfo(object networkSwitchObj)
        {
            if (networkSwitchObj == null)
                return null;

            string deviceName = DiscoverBestDeviceName(networkSwitchObj);
            string modelName = DiscoverBestModelName(networkSwitchObj);

            GameObject go = DiscoverGameObject(networkSwitchObj);
            Vector3 pos = go != null ? go.transform.position : Vector3.zero;

            return new RegisteredSwitchInfo
            {
                NetworkSwitchObject = networkSwitchObj,
                GameObject = go,
                DeviceName = deviceName,
                ModelName = modelName,
                WorldPosition = pos,
                HasWorldPosition = go != null
            };
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
            var groups = new List<List<RegisteredSwitchInfo>>();
            var remaining = new HashSet<RegisteredSwitchInfo>(switches);

            while (remaining.Count > 0)
            {
                RegisteredSwitchInfo seed = remaining.First();
                remaining.Remove(seed);

                var group = new List<RegisteredSwitchInfo>();
                var queue = new Queue<RegisteredSwitchInfo>();
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

        private void ApplySafeModeFabrics(List<List<RegisteredSwitchInfo>> fabrics)
        {
            FabricIdToSpeed.Clear();
            FabricIdToCableIds.Clear();
            CableIdToFabricId.Clear();
            CableIdToOverrideSpeed.Clear();

            int index = 1;

            foreach (List<RegisteredSwitchInfo> fabric in fabrics)
            {
                string fabricId = $"FABRIC-{index:000}";
                float fabricSpeed = fabric.Sum(x => x.EstimatedCapacityGbps);

                FabricIdToSpeed[fabricId] = fabricSpeed;
                FabricIdToCableIds[fabricId] = new HashSet<int>();

                string signature = BuildFabricSignature(fabric, fabricSpeed);
                if (!LastFabricSignatures.TryGetValue(fabricId, out string oldSig) ||
                    !string.Equals(oldSig, signature, StringComparison.Ordinal))
                {
                    LastFabricSignatures[fabricId] = signature;
                    LogToFile($"APPLY | fabricId={fabricId} | members={fabric.Count} | estCapacityGbps={fabricSpeed.ToString("0.##", CultureInfo.InvariantCulture)}");
                }

                HashSet<string> fabricDevices = new(
                    fabric.Select(x => x.DeviceName).Where(x => !string.IsNullOrWhiteSpace(x)),
                    StringComparer.OrdinalIgnoreCase);

                // 1. Internal switch-switch links only when BOTH ends are inside the same fabric.
                foreach (RegisteredCableInfo cable in RegisteredCables.Values)
                {
                    bool aIn = fabricDevices.Contains(cable.DeviceA);
                    bool bIn = fabricDevices.Contains(cable.DeviceB);

                    if (!(aIn && bIn))
                        continue;

                    FabricIdToCableIds[fabricId].Add(cable.CableId);
                    CableIdToFabricId[cable.CableId] = fabricId;
                    CableIdToOverrideSpeed[cable.CableId] = fabricSpeed;
                }

                // 2. Server-facing bundle links using NORMALIZED server keys.
                var serverBundles = new Dictionary<string, List<RegisteredCableInfo>>(StringComparer.OrdinalIgnoreCase);

                foreach (RegisteredCableInfo cable in RegisteredCables.Values)
                {
                    string serverKey = GetNormalizedServerBundleKey(cable);
                    if (string.IsNullOrWhiteSpace(serverKey))
                        continue;

                    bool switchAIn = fabricDevices.Contains(cable.DeviceA);
                    bool switchBIn = fabricDevices.Contains(cable.DeviceB);

                    // exactly one switch endpoint should be in the fabric for a server-facing cable
                    if (switchAIn == switchBIn)
                        continue;

                    if (!serverBundles.TryGetValue(serverKey, out List<RegisteredCableInfo> bundle))
                    {
                        bundle = new List<RegisteredCableInfo>();
                        serverBundles[serverKey] = bundle;
                    }

                    bundle.Add(cable);
                }

                foreach ((string serverKey, List<RegisteredCableInfo> bundle) in serverBundles)
                {
                    if (bundle.Count < 2)
                        continue;

                    float perCableSpeed = fabricSpeed / bundle.Count;

                    foreach (RegisteredCableInfo cable in bundle)
                    {
                        FabricIdToCableIds[fabricId].Add(cable.CableId);
                        CableIdToFabricId[cable.CableId] = fabricId;
                        CableIdToOverrideSpeed[cable.CableId] = perCableSpeed;
                    }

                    string cableList = string.Join(",", bundle.Select(x => x.CableId).OrderBy(x => x));
                    LogToFile($"SERVER BUNDLE | fabricId={fabricId} | serverKey={serverKey} | cableCount={bundle.Count} | perCableGbps={perCableSpeed.ToString("0.##", CultureInfo.InvariantCulture)} | cableIds=[{cableList}]");
                }

                foreach (RegisteredSwitchInfo sw in fabric)
                {
                    ApplyTagToSwitch(sw, fabricId, fabric.Count, fabricSpeed);
                }

                index++;
            }
        }

        private static string GetNormalizedServerBundleKey(RegisteredCableInfo cable)
        {
            string raw = GetServerEndpointName(cable);
            if (string.IsNullOrWhiteSpace(raw))
                return string.Empty;

            string normalized = raw.Trim();

            // Convert:
            // Server.Yellow2_-96608 -> Server.Yellow2
            // Server.Blue2_-1898036 -> Server.Blue2
            normalized = TrailingRuntimeIdRegex.Replace(normalized, "");

            return normalized;
        }

        private static string GetServerEndpointName(RegisteredCableInfo cable)
        {
            if (!string.IsNullOrWhiteSpace(cable.ExtraA) && cable.ExtraA.StartsWith("Server.", StringComparison.OrdinalIgnoreCase))
                return cable.ExtraA;

            if (!string.IsNullOrWhiteSpace(cable.ExtraB) && cable.ExtraB.StartsWith("Server.", StringComparison.OrdinalIgnoreCase))
                return cable.ExtraB;

            if (!string.IsNullOrWhiteSpace(cable.DeviceA) && cable.DeviceA.StartsWith("Server.", StringComparison.OrdinalIgnoreCase))
                return cable.DeviceA;

            if (!string.IsNullOrWhiteSpace(cable.DeviceB) && cable.DeviceB.StartsWith("Server.", StringComparison.OrdinalIgnoreCase))
                return cable.DeviceB;

            return string.Empty;
        }

        private static void ApplyTagToSwitch(RegisteredSwitchInfo sw, string fabricId, int memberCount, float speed)
        {
            if (sw?.GameObject == null)
                return;

            FabricGroupTag tag = sw.GameObject.GetComponent<FabricGroupTag>();
            if (tag == null)
                tag = sw.GameObject.AddComponent<FabricGroupTag>();

            tag.FabricId = fabricId;
            tag.MemberIndex = 0;
            tag.MemberCount = memberCount;
            tag.AggregatedBandwidth = speed;
            tag.AggregatedPortCount = 0;

            TrySetStringMember(sw.GameObject, "fabricGroupId", fabricId);
            TrySetStringMember(sw.GameObject, "lacpGroupId", fabricId);
            TrySetBoolMember(sw.GameObject, "fabricEnabled", true);
            TrySetFloatMember(sw.GameObject, "aggregatedBandwidth", speed);
            TrySetFloatMember(sw.GameObject, "fabricBandwidth", speed);
            TrySetFloatMember(sw.GameObject, "sharedBandwidth", speed);
            TrySetFloatMember(sw.GameObject, "maxBandwidth", speed);
            TrySetFloatMember(sw.GameObject, "throughput", speed);
            TrySetFloatMember(sw.GameObject, "maxThroughput", speed);
        }

        private static string BuildFabricSignature(List<RegisteredSwitchInfo> fabric, float speed)
        {
            string members = string.Join("|", fabric.OrderBy(x => x.DeviceName).Select(x => x.DeviceName));
            return $"{members}::{speed.ToString("0.###", CultureInfo.InvariantCulture)}";
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

        private static GameObject DiscoverGameObject(object obj)
        {
            if (obj == null)
                return null;

            try
            {
                if (obj is Component comp)
                    return comp.gameObject;
            }
            catch
            {
            }

            try
            {
                object go = TryGetMemberValue(obj, "gameObject");
                if (go is GameObject gameObject)
                    return gameObject;
            }
            catch
            {
            }

            try
            {
                object transform = TryGetMemberValue(obj, "transform");
                if (transform is Transform t)
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
                if (val is string s && !string.IsNullOrWhiteSpace(s))
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
                    if (f.GetValue(obj) is string s && !string.IsNullOrWhiteSpace(s))
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
                    if (p.GetValue(obj) is string s && !string.IsNullOrWhiteSpace(s))
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
            catch { return Array.Empty<FieldInfo>(); }
        }

        private static IEnumerable<PropertyInfo> SafeProperties(Type t)
        {
            try { return t.GetProperties(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance); }
            catch { return Array.Empty<PropertyInfo>(); }
        }

        private static List<int> ExtractIntList(object raw)
        {
            List<int> result = new();

            if (raw == null)
                return result;

            try
            {
                if (raw is int one)
                {
                    result.Add(one);
                    return result;
                }

                if (raw is IEnumerable enumerable && raw is not string)
                {
                    foreach (object item in enumerable)
                    {
                        if (item == null)
                            continue;

                        if (int.TryParse(item.ToString(), out int parsed))
                            result.Add(parsed);
                    }

                    return result.Distinct().ToList();
                }

                if (int.TryParse(raw.ToString(), out int single))
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
                    return prop.GetValue(target);

                FieldInfo field = t.GetField(memberName, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                if (field != null)
                    return field.GetValue(target);
            }
            catch
            {
            }

            return null;
        }

        private static void TrySetStringMember(object target, string memberName, string value) => TrySetMember(target, memberName, value);
        private static void TrySetBoolMember(object target, string memberName, bool value) => TrySetMember(target, memberName, value);
        private static void TrySetFloatMember(object target, string memberName, float value) => TrySetMember(target, memberName, value);

        private static void TrySetMember(object target, string memberName, object value)
        {
            if (target == null || string.IsNullOrWhiteSpace(memberName))
                return;

            try
            {
                Type t = target.GetType();

                FieldInfo field = t.GetField(memberName, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                if (field != null && field.FieldType.IsAssignableFrom(value.GetType()))
                {
                    field.SetValue(target, value);
                    return;
                }

                PropertyInfo prop = t.GetProperty(memberName, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                if (prop != null && prop.CanWrite && prop.PropertyType.IsAssignableFrom(value.GetType()))
                {
                    prop.SetValue(target, value);
                }
            }
            catch
            {
            }
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

        private static void LogToFile(string message)
        {
            try
            {
                Directory.CreateDirectory(DebugFolderPath);
                File.AppendAllText(DebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] {message}{Environment.NewLine}");
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
    }
}