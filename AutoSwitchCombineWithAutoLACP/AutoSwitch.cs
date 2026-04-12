using HarmonyLib;
using Il2CppInterop.Runtime.Injection;
using MelonLoader;
using MelonLoader.Utils;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using UnityEngine;
using Component = UnityEngine.Component;

[assembly: MelonInfo(typeof(AutoSwitch.AutoSwitchMod), "AutoSwitch", "0.1.3", "Big Texas Jerky")]
[assembly: MelonGame("Waseku", "Data Center")]

namespace AutoSwitch
{
    public sealed class AutoSwitchMod : MelonMod
    {
        private const float ScanIntervalSeconds = 8.0f;
        private const float SameRackXZTolerance = 0.20f;
        private const float AdjacentYTolerance = 0.030f;
        private const float FallbackExpectedStep = 0.0444f;

        private static readonly HashSet<string> AllowedSwitchNames = new(StringComparer.OrdinalIgnoreCase)
        {
            "Switch32xQSFP",
            "Switch4xQSXP16xSFP",
            "Switch4xSFP",
            "Switch16CU"
        };

        private static readonly string[] SwitchNameHints =
        {
            "switch",
            "qsfp",
            "sfp"
        };

        private static string DebugFolderPath =>
            Path.Combine(MelonEnvironment.ModsDirectory, "AutoSwitch");

        private static string DebugLogPath =>
            Path.Combine(DebugFolderPath, "autoswitch-debug.log");

        private float _nextScanTime;
        private string _lastSummary = string.Empty;
        private bool _loggedNearMissesThisScene;

        public override void OnInitializeMelon()
        {
            ClassInjector.RegisterTypeInIl2Cpp<FabricGroupTag>();
            Directory.CreateDirectory(DebugFolderPath);
            File.WriteAllText(DebugLogPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] AutoSwitch debug log started.{Environment.NewLine}");
            MelonLogger.Msg("[AutoSwitch] v0.1.3 debug build active. Loose detection mode.");
        }

        public override void OnSceneWasLoaded(int buildIndex, string sceneName)
        {
            if (buildIndex == 0)
                return;

            _nextScanTime = 0f;
            _lastSummary = string.Empty;
            _loggedNearMissesThisScene = false;
            LogToFile($"Scene loaded: {sceneName} ({buildIndex})");
        }

        public override void OnUpdate()
        {
            try
            {
                if (Time.time < _nextScanTime)
                    return;

                _nextScanTime = Time.time + ScanIntervalSeconds;
                RunDiscoveryScan();
            }
            catch (Exception ex)
            {
                MelonLogger.Error($"[AutoSwitch] OnUpdate exception: {ex}");
                LogToFile($"OnUpdate exception: {ex}");
                _nextScanTime = Time.time + 10.0f;
            }
        }

        private void RunDiscoveryScan()
        {
            List<SwitchNode> switches = DiscoverRealSwitchRoots();
            Dictionary<Transform, List<SwitchNode>> byRack = GroupByRackLikeParent(switches);

            int rackCount = byRack.Count;
            int switchCount = switches.Count;
            int adjacencyPairs = 0;

            foreach ((Transform rackRoot, List<SwitchNode> rackSwitches) in byRack)
            {
                List<SwitchNode> ordered = rackSwitches
                    .OrderByDescending(s => s.LocalPosition.y)
                    .ToList();

                for (int i = 1; i < ordered.Count; i++)
                {
                    if (AreAdjacent(ordered[i - 1], ordered[i], out _))
                        adjacencyPairs++;
                }
            }

            string summary =
                $"SCAN SUMMARY | racks={rackCount} | switches={switchCount} | adjacentPairs={adjacencyPairs}";

            if (!string.Equals(summary, _lastSummary, StringComparison.Ordinal))
            {
                _lastSummary = summary;
                MelonLogger.Msg($"[AutoSwitch] {summary}");
                LogToFile(summary);

                foreach ((Transform rackRoot, List<SwitchNode> rackSwitches) in byRack)
                {
                    List<SwitchNode> ordered = rackSwitches
                        .OrderByDescending(s => s.LocalPosition.y)
                        .ToList();

                    string rackName = rackRoot != null ? rackRoot.name : "<null>";
                    string rackLine =
                        $"GROUP | parent={rackName} | switchCount={ordered.Count} | switches={string.Join(", ", ordered.Select(s => $"{s.GameObject.name}[ports={s.TotalPorts}]"))}";
                    LogToFile(rackLine);
                }

                foreach (SwitchNode node in switches)
                {
                    LogToFile(
                        $"SWITCH | name={node.GameObject.name} | normalized={NormalizeCloneName(node.GameObject.name)} | parent={node.ParentName} | chain={node.ParentChain} | local=({node.LocalPosition.x:0.###},{node.LocalPosition.y:0.###},{node.LocalPosition.z:0.###}) | world=({node.WorldPosition.x:0.###},{node.WorldPosition.y:0.###},{node.WorldPosition.z:0.###}) | qsfp={node.QsfpPortCount} | sfp={node.SfpPortCount} | total={node.TotalPorts} | height={node.ApproxHeight:0.###}"
                    );
                }
            }

            if (!_loggedNearMissesThisScene)
            {
                _loggedNearMissesThisScene = true;
                DumpNearMisses();
            }
        }

        private static Dictionary<Transform, List<SwitchNode>> GroupByRackLikeParent(List<SwitchNode> switches)
        {
            var result = new Dictionary<Transform, List<SwitchNode>>();

            foreach (SwitchNode sw in switches)
            {
                Transform key = sw.GroupRoot ?? sw.GameObject.transform.parent;

                if (key == null)
                    continue;

                if (!result.TryGetValue(key, out List<SwitchNode> list))
                {
                    list = new List<SwitchNode>();
                    result[key] = list;
                }

                list.Add(sw);
            }

            return result;
        }

        private List<SwitchNode> DiscoverRealSwitchRoots()
        {
            var results = new List<SwitchNode>();
            Transform[] allTransforms = UnityEngine.Object.FindObjectsOfType<Transform>();
            HashSet<int> seen = new();

            foreach (Transform t in allTransforms)
            {
                if (t == null || t.gameObject == null)
                    continue;

                GameObject go = t.gameObject;
                int id = go.GetInstanceID();

                if (seen.Contains(id))
                    continue;

                if (!LooksLikeRealSwitchRoot(go, out int qsfpCount, out int sfpCount))
                    continue;

                Transform groupRoot = FindGroupingParent(go.transform);

                SwitchNode node = new SwitchNode
                {
                    GameObject = go,
                    GroupRoot = groupRoot,
                    ParentName = go.transform.parent != null ? go.transform.parent.name : "<root>",
                    ParentChain = BuildParentChain(go.transform, 5),
                    WorldPosition = go.transform.position,
                    LocalPosition = go.transform.localPosition,
                    ApproxHeight = EstimateObjectHeight(go),
                    QsfpPortCount = qsfpCount,
                    SfpPortCount = sfpCount
                };

                if (node.TotalPorts <= 0)
                    continue;

                results.Add(node);
                seen.Add(id);
            }

            return results;
        }

        private static bool LooksLikeRealSwitchRoot(GameObject go, out int qsfpCount, out int sfpCount)
        {
            qsfpCount = 0;
            sfpCount = 0;

            if (go == null)
                return false;

            string baseName = NormalizeCloneName(go.name);
            string lowerBase = baseName.ToLowerInvariant();

            bool allowedByExactName = AllowedSwitchNames.Contains(baseName);
            bool allowedByHint = SwitchNameHints.Any(h => lowerBase.Contains(h));

            if (!allowedByExactName && !allowedByHint)
                return false;

            qsfpCount = CountPorts(go, "QSFP_port.");
            sfpCount = CountPorts(go, "SFP_port.");

            int totalPorts = qsfpCount + sfpCount;

            if (allowedByExactName)
            {
                if (totalPorts < 4)
                    return false;
            }
            else
            {
                if (totalPorts < 8)
                    return false;
            }

            Transform parent = go.transform.parent;
            if (parent != null && parent.gameObject != null)
            {
                string parentBaseName = NormalizeCloneName(parent.gameObject.name);
                if (AllowedSwitchNames.Contains(parentBaseName))
                    return false;
            }

            return true;
        }

        private static Transform FindGroupingParent(Transform start)
        {
            if (start == null)
                return null;

            Transform current = start.parent;
            while (current != null)
            {
                string lower = NormalizeCloneName(current.name).ToLowerInvariant();

                if (lower.Contains("rack") ||
                    lower.Contains("usableobjects") ||
                    lower.Contains("customerbase") ||
                    lower.Contains("customerbases"))
                {
                    return current;
                }

                current = current.parent;
            }

            return start.parent;
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

        private static int CountPorts(GameObject go, string prefix)
        {
            int count = 0;

            foreach (Transform t in go.GetComponentsInChildren<Transform>(true))
            {
                if (t == null)
                    continue;

                string n = t.name ?? string.Empty;
                if (n.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    count++;
            }

            return count;
        }

        private static float EstimateObjectHeight(GameObject go)
        {
            if (go == null)
                return FallbackExpectedStep;

            Renderer[] renderers = go.GetComponentsInChildren<Renderer>(true);
            if (renderers == null || renderers.Length == 0)
                return FallbackExpectedStep;

            Bounds bounds = renderers[0].bounds;
            for (int i = 1; i < renderers.Length; i++)
                bounds.Encapsulate(renderers[i].bounds);

            return Mathf.Max(0.001f, bounds.size.y);
        }

        private static bool AreAdjacent(SwitchNode upper, SwitchNode lower, out string reason)
        {
            float dx = Mathf.Abs(upper.LocalPosition.x - lower.LocalPosition.x);
            float dz = Mathf.Abs(upper.LocalPosition.z - lower.LocalPosition.z);

            if (dx > SameRackXZTolerance || dz > SameRackXZTolerance)
            {
                reason = $"XZ too far | dx={dx:0.###} dz={dz:0.###}";
                return false;
            }

            float expectedStep = Mathf.Max(
                upper.ApproxHeight > 0.001f ? upper.ApproxHeight : FallbackExpectedStep,
                lower.ApproxHeight > 0.001f ? lower.ApproxHeight : FallbackExpectedStep
            );

            float actualStep = Mathf.Abs(upper.LocalPosition.y - lower.LocalPosition.y);

            if (Mathf.Abs(actualStep - expectedStep) > AdjacentYTolerance)
            {
                reason = $"Y gap mismatch | actual={actualStep:0.###} expected={expectedStep:0.###}";
                return false;
            }

            reason = "OK";
            return true;
        }

        private void DumpNearMisses()
        {
            try
            {
                Transform[] allTransforms = UnityEngine.Object.FindObjectsOfType<Transform>();

                LogToFile("NEAR MISS DUMP START");

                int written = 0;
                foreach (Transform t in allTransforms)
                {
                    if (t == null || t.gameObject == null)
                        continue;

                    string name = t.gameObject.name ?? string.Empty;
                    string lower = name.ToLowerInvariant();

                    if (!(lower.Contains("switch") || lower.Contains("qsfp") || lower.Contains("sfp")))
                        continue;

                    int qsfpCount = CountPorts(t.gameObject, "QSFP_port.");
                    int sfpCount = CountPorts(t.gameObject, "SFP_port.");

                    LogToFile(
                        $"NEAR MISS | name={t.gameObject.name} | normalized={NormalizeCloneName(t.gameObject.name)} | parent={t.parent?.name ?? "<root>"} | chain={BuildParentChain(t, 4)} | qsfp={qsfpCount} | sfp={sfpCount}"
                    );

                    written++;
                    if (written >= 60)
                        break;
                }

                LogToFile("NEAR MISS DUMP END");
            }
            catch (Exception ex)
            {
                LogToFile($"DumpNearMisses failed: {ex}");
            }
        }

        private static string BuildParentChain(Transform t, int depth)
        {
            if (t == null)
                return "<null>";

            List<string> parts = new();
            Transform current = t.parent;
            int count = 0;

            while (current != null && count < depth)
            {
                parts.Add(current.name);
                current = current.parent;
                count++;
            }

            if (parts.Count == 0)
                return "<root>";

            return string.Join(" <- ", parts);
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

    internal sealed class SwitchNode
    {
        public GameObject GameObject;
        public Transform GroupRoot;
        public string ParentName;
        public string ParentChain;
        public Vector3 WorldPosition;
        public Vector3 LocalPosition;
        public float ApproxHeight;
        public int QsfpPortCount;
        public int SfpPortCount;

        public int TotalPorts => QsfpPortCount + SfpPortCount;
    }
}