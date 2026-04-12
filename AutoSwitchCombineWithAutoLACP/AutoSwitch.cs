using HarmonyLib;
using Il2CppInterop.Runtime.Injection;
using MelonLoader;
using MelonLoader.Utils;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using UnityEngine;
using Component = UnityEngine.Component;

[assembly: MelonInfo(typeof(AutoSwitch.AutoSwitchMod), "AutoSwitch", "0.1.2", "Big Texas Jerky")]
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

        private static readonly HashSet<string> AllowedRackNames = new(StringComparer.OrdinalIgnoreCase)
        {
            "Rack_Lanberg_47U",
            "BoxedRack"
        };

        private static string DebugFolderPath =>
            Path.Combine(MelonEnvironment.ModsDirectory, "AutoSwitch");

        private static string DebugLogPath =>
            Path.Combine(DebugFolderPath, "autoswitch-debug.log");

        private float _nextScanTime;
        private string _lastSummary = string.Empty;

        public override void OnInitializeMelon()
        {
            ClassInjector.RegisterTypeInIl2Cpp<FabricGroupTag>();
            Directory.CreateDirectory(DebugFolderPath);
            File.WriteAllText(DebugLogPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] AutoSwitch debug log started.{Environment.NewLine}");
            MelonLogger.Msg("[AutoSwitch] Debug build active. Automatic scan logging only.");
        }

        public override void OnSceneWasLoaded(int buildIndex, string sceneName)
        {
            if (buildIndex == 0)
                return;

            _nextScanTime = 0f;
            _lastSummary = string.Empty;
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
            Dictionary<Transform, List<SwitchNode>> byRack = GroupByRack(switches);

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

                    string rackLine =
                        $"RACK | name={rackRoot.name} | switchCount={ordered.Count} | switches={string.Join(", ", ordered.Select(s => $"{s.GameObject.name}[ports={s.TotalPorts}]"))}";
                    LogToFile(rackLine);
                }
            }
        }

        private static Dictionary<Transform, List<SwitchNode>> GroupByRack(List<SwitchNode> switches)
        {
            var result = new Dictionary<Transform, List<SwitchNode>>();

            foreach (SwitchNode sw in switches)
            {
                if (sw.RackRoot == null)
                    continue;

                if (!result.TryGetValue(sw.RackRoot, out List<SwitchNode> list))
                {
                    list = new List<SwitchNode>();
                    result[sw.RackRoot] = list;
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

                if (!LooksLikeRealSwitchRoot(go))
                    continue;

                Transform rackRoot = FindAllowedRackRoot(go.transform);
                if (rackRoot == null)
                    continue;

                SwitchNode node = new SwitchNode
                {
                    GameObject = go,
                    RackRoot = rackRoot,
                    WorldPosition = go.transform.position,
                    LocalPosition = go.transform.localPosition,
                    ApproxHeight = EstimateObjectHeight(go),
                    QsfpPortCount = CountPorts(go, "QSFP_port."),
                    SfpPortCount = CountPorts(go, "SFP_port.")
                };

                if (node.TotalPorts <= 0)
                    continue;

                results.Add(node);
                seen.Add(id);
            }

            return results;
        }

        private static bool LooksLikeRealSwitchRoot(GameObject go)
        {
            if (go == null)
                return false;

            string baseName = NormalizeCloneName(go.name);
            if (!AllowedSwitchNames.Contains(baseName))
                return false;

            int qsfpCount = CountPorts(go, "QSFP_port.");
            int sfpCount = CountPorts(go, "SFP_port.");
            int totalPorts = qsfpCount + sfpCount;

            if (totalPorts < 4)
                return false;

            Transform parent = go.transform.parent;
            if (parent != null && parent.gameObject != null)
            {
                string parentBaseName = NormalizeCloneName(parent.gameObject.name);
                if (AllowedSwitchNames.Contains(parentBaseName))
                    return false;
            }

            return true;
        }

        private static Transform FindAllowedRackRoot(Transform start)
        {
            if (start == null)
                return null;

            Transform current = start.parent;
            while (current != null)
            {
                string baseName = NormalizeCloneName(current.name);
                if (AllowedRackNames.Contains(baseName))
                    return current;

                current = current.parent;
            }

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
        public Transform RackRoot;
        public Vector3 WorldPosition;
        public Vector3 LocalPosition;
        public float ApproxHeight;
        public int QsfpPortCount;
        public int SfpPortCount;

        public int TotalPorts => QsfpPortCount + SfpPortCount;
    }
}