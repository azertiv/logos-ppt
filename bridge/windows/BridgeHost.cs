using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Microsoft.Win32.SafeHandles;

namespace Neurow.Pictos
{
    internal sealed class BridgeHost : IDisposable
    {
        private Process process;
        private SafeJobHandle job;
        internal event Action<string> Output;
        internal event Action Exited;
        internal bool Running => process != null && !process.HasExited;
        internal Process Process => process;
        internal void Start(string root, string directory, string token)
        {
            if (process != null) throw new InvalidOperationException("Already started.");
            string node = Path.Combine(root, "runtime", "node.exe");
            string server = Path.Combine(root, "bridge", "server.js");
            if (!File.Exists(node) || !File.Exists(server)) throw new FileNotFoundException("Décompressez le dossier entier du compagnon avant de le lancer.");
            var start = new ProcessStartInfo(node, "\"" + server + "\" --managed")
            {
                WorkingDirectory = root, UseShellExecute = false, CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden, RedirectStandardInput = true,
                RedirectStandardOutput = true, RedirectStandardError = true
            };
            start.EnvironmentVariables["PICTOS_DATA_DIR"] = directory;
            start.EnvironmentVariables["PICTOS_PORT"] = "43129";
            // A job owns Node and every Codex descendant, including on a launcher crash.
            job = Job.Create();
            process = new Process { StartInfo = start, EnableRaisingEvents = true };
            process.OutputDataReceived += (_, e) => { if (e.Data != null && e.Data.Length < 4096) Output?.Invoke(e.Data); };
            process.ErrorDataReceived += (_, e) => { if (e.Data != null && e.Data.Length < 4096) Output?.Invoke(e.Data); };
            process.Exited += (_, __) => Exited?.Invoke();
            try
            {
                process.Start();
                if (!Job.AssignProcessToJobObject(job, process.Handle)) throw new Win32Exception(Marshal.GetLastWin32Error());
                process.BeginOutputReadLine(); process.BeginErrorReadLine();
                process.StandardInput.WriteLine("{\"pairingToken\":\"" + token + "\"}");
                process.StandardInput.Flush();
            }
            catch { try { if (Running) process.Kill(); } catch { } Dispose(); throw; }
        }
        internal async Task StopAsync()
        {
            if (Running)
            {
                try { process.StandardInput.WriteLine("{\"command\":\"stop\"}"); process.StandardInput.Flush(); } catch { }
                await Task.Run(() => { try { process.WaitForExit(3000); } catch { } });
            }
            Dispose();
        }
        public void Dispose()
        {
            job?.Dispose(); job = null;
            process?.Dispose(); process = null;
        }
    }
    internal sealed class SafeJobHandle : SafeHandleZeroOrMinusOneIsInvalid
    {
        private SafeJobHandle() : base(true) { }
        protected override bool ReleaseHandle() { return Job.CloseHandle(handle); }
    }
    internal static class Job
    {
        [StructLayout(LayoutKind.Sequential)] private struct BasicLimits
        {
            internal long PerProcessUserTimeLimit, PerJobUserTimeLimit;
            internal uint LimitFlags;
            internal UIntPtr MinimumWorkingSetSize, MaximumWorkingSetSize;
            internal uint ActiveProcessLimit;
            internal UIntPtr Affinity;
            internal uint PriorityClass, SchedulingClass;
        }
        [StructLayout(LayoutKind.Sequential)] private struct IoCounters
        { internal ulong ReadOperationCount, WriteOperationCount, OtherOperationCount, ReadTransferCount, WriteTransferCount, OtherTransferCount; }
        [StructLayout(LayoutKind.Sequential)] private struct ExtendedLimits
        {
            internal BasicLimits BasicLimitInformation;
            internal IoCounters IoInfo;
            internal UIntPtr ProcessMemoryLimit, JobMemoryLimit, PeakProcessMemoryUsed, PeakJobMemoryUsed;
        }
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)] private static extern SafeJobHandle CreateJobObject(IntPtr attributes, string name);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern bool SetInformationJobObject(SafeJobHandle job, int infoClass, ref ExtendedLimits info, uint length);
        [DllImport("kernel32.dll", SetLastError = true)] internal static extern bool AssignProcessToJobObject(SafeJobHandle job, IntPtr process);
        [DllImport("kernel32.dll")] internal static extern bool CloseHandle(IntPtr handle);
        internal static SafeJobHandle Create()
        {
            var job = CreateJobObject(IntPtr.Zero, null);
            if (job.IsInvalid) throw new Win32Exception(Marshal.GetLastWin32Error());
            var limits = new ExtendedLimits { BasicLimitInformation = new BasicLimits { LimitFlags = 0x2000 } };
            if (!SetInformationJobObject(job, 9, ref limits, (uint)Marshal.SizeOf(limits)))
            { int error = Marshal.GetLastWin32Error(); job.Dispose(); throw new Win32Exception(error); }
            return job;
        }
    }
}
