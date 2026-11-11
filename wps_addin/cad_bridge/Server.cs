using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace Dongsidaye.WpsCad
{
    internal sealed class ReadyMarker
    {
        public string exe { get; set; }
        public string request { get; set; }
        public string workerHash { get; set; }
    }

    internal sealed class ServerHeartbeat
    {
        public int protocol { get; set; }
        public string exe { get; set; }
        public string instance { get; set; }
        public long heartbeatAt { get; set; }
        public long startedAt { get; set; }
        public long startedTick { get; set; }
        public long processStartedAt { get; set; }
        public int pid { get; set; }
        public string workerHash { get; set; }
    }

    internal sealed class OwnedPhase : IDisposable
    {
        internal readonly JobContext Job;
        internal readonly Process Process;
        internal readonly PhaseProcessGroup Group;
        internal readonly Stopwatch Elapsed = Stopwatch.StartNew();
        internal bool TimedOut;

        internal OwnedPhase(JobContext job, Process process, PhaseProcessGroup group)
        { Job = job; Process = process; Group = group; }

        public void Dispose()
        {
            Group.Dispose();
            Process.Dispose();
        }
    }

    internal sealed class CadServer : IDisposable
    {
        private readonly string executable = SafeFiles.ExistingLongPath(Assembly.GetExecutingAssembly().Location);
        private readonly string binaryHash = SafeFiles.Hash(SafeFiles.Read(Assembly.GetExecutingAssembly().Location, SafeFiles.MaximumBytes));
        private readonly string temp = SafeFiles.ExistingLongPath(Path.GetTempPath());
        private readonly string instance = Guid.NewGuid().ToString("N");
        private readonly long startedAt = Now();
        private readonly long startedTick = Stopwatch.GetTimestamp();
        private readonly long processStartedAt = Unix(Process.GetCurrentProcess().StartTime.ToUniversalTime());
        private readonly int pid = Process.GetCurrentProcess().Id;
        private readonly List<OwnedPhase> phases = new List<OwnedPhase>();
        private readonly Dictionary<string, string> failedStarts = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private readonly object heartbeatLock = new object();
        private readonly Stopwatch idle = Stopwatch.StartNew();
        private Timer heartbeatTimer;
        private volatile bool stopping;
        private volatile bool retiring;
        private string HeartbeatPath { get { return Path.Combine(temp, "DongsidayeWpsCad." + binaryHash + ".server.json"); } }

        internal static int Run()
        {
            try
            {
                using (var server = new CadServer())
                {
                    string identity = WindowsIdentity.GetCurrent().User.Value + "|" + server.executable.ToUpperInvariant() + "|" + server.binaryHash;
                    string mutexName = "Global\\DongsidayeWpsCadServer_" + SafeFiles.Hash(Encoding.UTF8.GetBytes(identity));
                    using (var mutex = new Mutex(false, mutexName))
                    {
                        bool owned;
                        try { owned = mutex.WaitOne(0); }
                        catch (AbandonedMutexException) { owned = true; }
                        if (!owned) return 0;
                        try { return server.Loop(); }
                        finally { mutex.ReleaseMutex(); }
                    }
                }
            }
            catch { return 1; }
        }

        private int Loop()
        {
            WriteHeartbeat();
            heartbeatTimer = new Timer(delegate { WriteHeartbeat(); }, null, 2000, 2000);
            while (!stopping)
            {
                CheckRetirement();
                ReapPhases();
                if (retiring && phases.Count == 0) break;
                if (!retiring) QueuePhases();
                if (phases.Count == 0 && idle.Elapsed.TotalMinutes >= 60) break;
                Thread.Sleep(150);
            }
            return 0;
        }

        private void QueuePhases()
        {
            foreach (string finished in failedStarts.Keys.Where(path => !File.Exists(path)).ToArray()) failedStarts.Remove(finished);
            IEnumerable<string> candidates;
            try
            {
                candidates = Directory.EnumerateFiles(temp, "DongsidayeWpsCad_*.ready", SearchOption.TopDirectoryOnly)
                    .OrderBy(path => File.GetLastWriteTimeUtc(path)).ThenBy(path => path, StringComparer.OrdinalIgnoreCase).ToArray();
            }
            catch { return; }
            foreach (string path in candidates)
            {
                if (phases.Count >= 2 || retiring || stopping) return;
                Match match = Regex.Match(Path.GetFileName(path), @"^DongsidayeWpsCad_([a-f0-9]{32})\.(convert|prepare|finalize|cleanup)\.ready$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                if (!match.Success || phases.Any(phase => phase.Job.Id == match.Groups[1].Value.ToLowerInvariant())) continue;
                try
                {
                    SafeFiles.RejectReparse(path);
                    ReadyMarker marker = SafeFiles.ReadJson<ReadyMarker>(path, 65536);
                    if (marker == null || marker.workerHash != binaryHash || String.IsNullOrEmpty(marker.exe) || String.IsNullOrEmpty(marker.request) ||
                        !String.Equals(SafeFiles.ExistingLongPath(marker.exe), executable, StringComparison.OrdinalIgnoreCase)) continue;
                    var job = new JobContext(marker.request);
                    if (job.Id != match.Groups[1].Value.ToLowerInvariant() || job.Phase != match.Groups[2].Value.ToLowerInvariant()) continue;
                    string expectedRequest = path.Substring(0, path.Length - ".ready".Length) + ".json";
                    if (!String.Equals(SafeFiles.ExistingLongPath(job.RequestPath), SafeFiles.ExistingLongPath(expectedRequest), StringComparison.OrdinalIgnoreCase)) continue;
                    if (HasTerminalStatus(job)) { File.Delete(path); continue; }
                    if (job.Phase != "cleanup" && File.Exists(job.CancelPath))
                    {
                        CleanupCancelledJob(job);
                        job.Complete(false, "", "已取消 CAD 导入。", true);
                        File.Delete(path);
                        continue;
                    }
                    string failure;
                    if (failedStarts.TryGetValue(path, out failure))
                    {
                        // Retry only the terminal notification if its file was
                        // temporarily locked. Never spawn this failed start again.
                        job.Complete(false, "", failure, false);
                        File.Delete(path);
                        continue;
                    }
                    OwnedPhase phase;
                    try { phase = StartPhase(job); }
                    catch (Exception error)
                    {
                        // All target paths are validated above. Native startup
                        // failures need a terminal result instead of a spawn loop.
                        failure = "无法启动 CAD 处理阶段：" + error.Message;
                        failedStarts[path] = failure;
                        job.Complete(false, "", failure, false);
                        File.Delete(path);
                        continue;
                    }
                    phases.Add(phase);
                    idle.Restart();
                    File.Delete(path);
                }
                catch { }
            }
        }

        private OwnedPhase StartPhase(JobContext job)
        {
            Process process = null;
            PhaseProcessGroup group = null;
            NativeProcessInfo native = new NativeProcessInfo();
            try
            {
                if (job.RequestPath.IndexOf('"') >= 0) throw new InvalidOperationException("CAD 请求路径无效。");
                group = new PhaseProcessGroup();
                var startup = new StartupInfo { cb = Marshal.SizeOf(typeof(StartupInfo)), dwFlags = 1, wShowWindow = 0 };
                var command = new StringBuilder("\"" + executable + "\" --job \"" + job.RequestPath + "\"");
                // No worker code executes before it belongs to the server's
                // kill-on-close job, so a converter cannot escape supervision.
                if (!CreateProcess(executable, command, IntPtr.Zero, IntPtr.Zero, false, 0x08000004,
                    IntPtr.Zero, Path.GetDirectoryName(executable), ref startup, out native))
                    throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
                group.Assign(native.Process);
                process = Process.GetProcessById((int)native.ProcessId);
                if (process.Handle == IntPtr.Zero) throw new InvalidOperationException("无法取得 CAD 子进程句柄。");
                if (ResumeThread(native.Thread) == UInt32.MaxValue) throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
                return new OwnedPhase(job, process, group);
            }
            catch
            {
                if (group != null) group.Dispose();
                if (native.Process != IntPtr.Zero)
                {
                    // This is the exact handle returned by our suspended spawn.
                    TerminateProcess(native.Process, 1);
                    WaitForSingleObject(native.Process, 5000);
                }
                if (process != null) process.Dispose();
                throw;
            }
            finally
            {
                if (native.Thread != IntPtr.Zero) CloseNativeHandle(native.Thread);
                if (native.Process != IntPtr.Zero) CloseNativeHandle(native.Process);
            }
        }

        private static void CleanupCancelledJob(JobContext job)
        {
            try
            {
                SafeFiles.RejectReparse(job.Prefix + ".lock");
                using (var claimed = new FileStream(job.Prefix + ".lock", FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.Read))
                    SafeFiles.DeleteFolder(job);
            }
            catch { }
        }

        private void ReapPhases()
        {
            for (int index = phases.Count - 1; index >= 0; index--)
            {
                OwnedPhase phase = phases[index];
                bool exited;
                try { exited = phase.Process.HasExited; }
                catch { exited = true; }
                if (!exited && phase.Elapsed.Elapsed.TotalSeconds > 240)
                {
                    phase.TimedOut = true;
                    try { SafeFiles.WriteText(phase.Job.CancelPath, "1"); } catch { }
                    phase.Group.Terminate();
                    try { exited = phase.Process.WaitForExit(5000); } catch { }
                }
                if (!exited) continue;
                int? exitCode = null;
                try { exitCode = phase.Process.ExitCode; } catch { }
                phase.Dispose();
                phases.RemoveAt(index);
                idle.Restart();
                // The frontend acknowledges cleanup by removing its terminal
                // status. A successful exit must not recreate that acknowledged
                // file as a failure during the server's later reap turn.
                bool cleanupAcknowledged = phase.Job.Phase == "cleanup" && exitCode == 0;
                if (!HasTerminalStatus(phase.Job) && !cleanupAcknowledged)
                {
                    try { phase.Job.Complete(false, "", phase.TimedOut ? "CAD 阶段超过时间限制，已停止。" : "CAD 处理组件异常结束，请重新导入。", false); } catch { }
                }
                if (phase.TimedOut)
                {
                    try { SafeFiles.DeleteFolder(phase.Job); } catch { }
                }
            }
        }

        private static bool HasTerminalStatus(JobContext job)
        {
            try
            {
                var status = SafeFiles.ReadJson<Dictionary<string, object>>(job.StatusPath, 65536);
                return status != null && status.ContainsKey("state") && String.Equals(System.Convert.ToString(status["state"]), "done", StringComparison.Ordinal);
            }
            catch { return false; }
        }

        private ServerHeartbeat ReadHeartbeat()
        {
            try { return SafeFiles.ReadJson<ServerHeartbeat>(HeartbeatPath, 65536); }
            catch { return null; }
        }

        private bool NewerLiveInstance(ServerHeartbeat other)
        {
            if (other == null || other.protocol != 1 || other.workerHash != binaryHash || other.instance == instance || other.pid <= 0) return false;
            if (other.startedTick < startedTick || (other.startedTick == startedTick && String.CompareOrdinal(other.instance, instance) <= 0)) return false;
            long age = Now() - other.heartbeatAt;
            if (age < -1000 || age > 8000) return false;
            try
            {
                using (Process process = Process.GetProcessById(other.pid))
                {
                    return !process.HasExited && Math.Abs(Unix(process.StartTime.ToUniversalTime()) - other.processStartedAt) < 1000 &&
                        String.Equals(SafeFiles.ExistingLongPath(process.MainModule.FileName), SafeFiles.ExistingLongPath(other.exe), StringComparison.OrdinalIgnoreCase);
                }
            }
            catch { return false; }
        }

        private void CheckRetirement()
        {
            if (!retiring && NewerLiveInstance(ReadHeartbeat())) retiring = true;
        }

        private void WriteHeartbeat()
        {
            lock (heartbeatLock)
            {
                if (stopping || retiring) return;
                try
                {
                    CheckRetirement();
                    if (retiring) return;
                    SafeFiles.WriteJson(HeartbeatPath, new ServerHeartbeat { protocol = 1, exe = executable, instance = instance,
                        heartbeatAt = Now(), startedAt = startedAt, startedTick = startedTick, pid = pid, processStartedAt = processStartedAt, workerHash = binaryHash });
                }
                catch { }
            }
        }

        private static long Unix(DateTime utc) { return (long)(utc - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalMilliseconds; }
        private static long Now() { return Unix(DateTime.UtcNow); }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct StartupInfo
        {
            public int cb;
            public string lpReserved, lpDesktop, lpTitle;
            public uint dwX, dwY, dwXSize, dwYSize, dwXCountChars, dwYCountChars, dwFillAttribute, dwFlags;
            public short wShowWindow, cbReserved2;
            public IntPtr lpReserved2, hStdInput, hStdOutput, hStdError;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct NativeProcessInfo
        {
            public IntPtr Process, Thread;
            public uint ProcessId, ThreadId;
        }

        [DllImport("kernel32.dll", EntryPoint = "CreateProcessW", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool CreateProcess(string applicationName, StringBuilder commandLine, IntPtr processAttributes,
            IntPtr threadAttributes, bool inheritHandles, uint creationFlags, IntPtr environment,
            string currentDirectory, ref StartupInfo startupInfo, out NativeProcessInfo processInformation);
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern uint ResumeThread(IntPtr thread);
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool TerminateProcess(IntPtr process, uint exitCode);
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern uint WaitForSingleObject(IntPtr handle, uint milliseconds);
        [DllImport("kernel32.dll", EntryPoint = "CloseHandle", SetLastError = true)]
        private static extern bool CloseNativeHandle(IntPtr handle);

        public void Dispose()
        {
            stopping = true;
            if (heartbeatTimer != null) heartbeatTimer.Dispose();
            foreach (OwnedPhase phase in phases) phase.Dispose();
            phases.Clear();
            lock (heartbeatLock)
            {
                ServerHeartbeat current = ReadHeartbeat();
                if (current != null && current.instance == instance && current.pid == pid)
                {
                    try { SafeFiles.RejectReparse(HeartbeatPath); File.Delete(HeartbeatPath); } catch { }
                }
            }
        }
    }

    internal sealed class PhaseProcessGroup : IDisposable
    {
        private IntPtr handle;

        internal PhaseProcessGroup()
        {
            handle = CreateJobObject(IntPtr.Zero, null);
            if (handle == IntPtr.Zero) throw new System.ComponentModel.Win32Exception();
            var limits = new ExtendedLimits();
            limits.BasicLimitInformation.LimitFlags = 0x2000; // JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
            if (!SetInformationJobObject(handle, 9, ref limits, (uint)Marshal.SizeOf(typeof(ExtendedLimits))))
            {
                int error = Marshal.GetLastWin32Error();
                Dispose();
                throw new System.ComponentModel.Win32Exception(error);
            }
        }

        internal void Assign(IntPtr process)
        {
            if (!AssignProcessToJobObject(handle, process)) throw new System.ComponentModel.Win32Exception();
        }

        internal void Terminate()
        {
            if (handle != IntPtr.Zero) TerminateJobObject(handle, 1);
        }

        public void Dispose()
        {
            if (handle != IntPtr.Zero) { CloseHandle(handle); handle = IntPtr.Zero; }
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct BasicLimits
        {
            public long PerProcessUserTimeLimit, PerJobUserTimeLimit;
            public uint LimitFlags;
            public UIntPtr MinimumWorkingSetSize, MaximumWorkingSetSize;
            public uint ActiveProcessLimit;
            public UIntPtr Affinity;
            public uint PriorityClass, SchedulingClass;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct IoCounters
        {
            public ulong ReadOperationCount, WriteOperationCount, OtherOperationCount;
            public ulong ReadTransferCount, WriteTransferCount, OtherTransferCount;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct ExtendedLimits
        {
            public BasicLimits BasicLimitInformation;
            public IoCounters IoInfo;
            public UIntPtr ProcessMemoryLimit, JobMemoryLimit, PeakProcessMemoryUsed, PeakJobMemoryUsed;
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr CreateJobObject(IntPtr attributes, string name);
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool SetInformationJobObject(IntPtr job, int informationClass, ref ExtendedLimits information, uint length);
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool AssignProcessToJobObject(IntPtr job, IntPtr process);
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool TerminateJobObject(IntPtr job, uint exitCode);
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool CloseHandle(IntPtr value);
    }
}
