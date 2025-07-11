using Spectre.Console;
using System;
using System.Diagnostics;
using LibreHardwareMonitor.Hardware;
using System.Management;

public static class Program
{
    public record ProcessInfo(string ProcessName, int ProcessId);
    public static void Main(string[] args)
    {
        #region Show ReadMe Information
        var objheaderLine = new Rule("[bold blue]Initiating Energy Consumption Evaluation Procedure[/]");
        AnsiConsole.Write(objheaderLine);

        var panel = new Panel("sdfhsnsfdsssssssssssssssssssss");
        panel.Header = new PanelHeader("Some text");
        panel.Expand = true;
        #endregion

        #region Initiating process
        //AnsiConsole.Write(new Markup("[/]"));
        AnsiConsole.Progress()
        .Start(ctx =>
        {
            var task1 = ctx.AddTask("[green]Fetching list of active processes[/]");
            var task2 = ctx.AddTask("[green]Fetching process IDs[/]");
            var task3 = ctx.AddTask("[green]Generating process list[/]");

            while (!ctx.IsFinished)
            {
                task1.Increment(1.5);
                task2.Increment(0.5);
                task3.Increment(1);
            }
        });
        #endregion

        #region Get list of active processes
        var lstProcesses = new List<string>(
            );
        var numProcessCount = 1;
        var lstProcessData = new List<ProcessInfo>();
        foreach (var p in Process.GetProcesses().OrderBy(p => p.ProcessName))
        {
            lstProcessData.Add(new ProcessInfo(p.ProcessName, p.Id));
            lstProcesses.Add($"{numProcessCount++}. {p.ProcessName} (Process Id: {p.Id})");
        }
        #endregion

        #region Display list of active processes
        var table = new Table().LeftAligned();
        numProcessCount = 1;
        AnsiConsole.Live(table)
            .Overflow(VerticalOverflow.Ellipsis) // Show ellipsis when overflowing
            .Cropping(VerticalOverflowCropping.Bottom) // Crop overflow at top
            .Start(ctx =>
            {
                table.Border(TableBorder.HeavyEdge);

                table.AddColumn("#");
                ctx.Refresh();
                Thread.Sleep(100);

                table.AddColumn("Process Name");
                ctx.Refresh();
                Thread.Sleep(100);

                table.AddColumn("Process ID");
                ctx.Refresh();
                Thread.Sleep(100);

                foreach (var process in lstProcessData)
                {
                    table.AddRow([Convert.ToString(numProcessCount++), Convert.ToString(process.ProcessName), Convert.ToString(process.ProcessId)]);
                }
            });
        #endregion

        #region Get process to monitor
        var strProcess = AnsiConsole.Prompt(
        new SelectionPrompt<string>()
        .Title("[blue] Select the process to monitor [/]:")
        .PageSize(25)
        .MoreChoicesText("[grey](Move up and down to reveal more choices)[/]")
        .AddChoices(lstProcesses));
        #endregion

        #region Extract process Id
        int numProcessId = ExtractProcessId(strProcess);
        #endregion

        #region Monitor Process
        var numDurationToMonitor = AnsiConsole.Prompt(
            new TextPrompt<int>("Enter duration to monitor (seconds): ")
            .Validate((duration) => duration switch
            {
                < 10 => ValidationResult.Error("Invalid Duration"),
                >= 10 => ValidationResult.Success(),
            }
        ));

        AnsiConsole.Progress()
        .Columns(
            [
                new TaskDescriptionColumn(),    // Task description
                new ProgressBarColumn(),        // Progress bar
                new PercentageColumn(),         // Percentage
                new RemainingTimeColumn(),      // Remaining time
                new SpinnerColumn(),            // Spinner
                new DownloadedColumn(),         // Downloaded
            ])
            .StartAsync(async ctx =>
            {
                var task = ctx.AddTask("[green]Monitoring[/]", new ProgressTaskSettings());
                for (int i = numDurationToMonitor; i >= 0; i--)
                {
                    task.Increment(100 / numDurationToMonitor);
                    await Task.Delay(1000);
                }
            });

        AnsiConsole.Progress()
        .Start(async ctx =>
        {
            var task1 = ctx.AddTask("[green]Generating final report...[/]");

            while (!ctx.IsFinished)
            {
                task1.Increment(2);
            }
        });

        Process objProcess = GetOrStartProcessById(numProcessId);
        string instanceName = GetProcessInstanceName(objProcess.Id);

        PerformanceCounter diskIO = new PerformanceCounter("Process", "IO Data Bytes/sec", instanceName);
        PerformanceCounter diskQueue = new PerformanceCounter("PhysicalDisk", "Avg. Disk Queue Length", "_Total");

        _ = diskIO.NextValue(); // reset

        TimeSpan startCpu = objProcess.TotalProcessorTime;
        DateTime start = DateTime.Now;

        Thread.Sleep(numDurationToMonitor * 1000);
        objProcess.Refresh();

        TimeSpan endCpu = objProcess.TotalProcessorTime;
        double cpuMs = (endCpu - startCpu).TotalMilliseconds;
        double cpuUsage = cpuMs / (Environment.ProcessorCount * numDurationToMonitor * 1000) * 100;
        double memoryMB = objProcess.WorkingSet64 / (1024.0 * 1024.0);
        double diskKBs = diskIO.NextValue() / 1024.0;
        float diskQueueLen = diskQueue.NextValue();

        double powerDrawWatts = 30;
        double hours = numDurationToMonitor / 3600.0;
        double energyKWh = (cpuUsage / 100.0) * powerDrawWatts * hours;
        double emissionFactor = 708;
        double co2Grams = energyKWh * emissionFactor;

        //Hardware Info
        var computer = new Computer
        {
            IsCpuEnabled = true,
            IsGpuEnabled = true,
            IsBatteryEnabled = true
        };
        computer.Open();

        float? cpuTemp = null, cpuClock = null, gpuTemp = null, gpuLoad = null;

        foreach (IHardware hardware in computer.Hardware)
        {
            hardware.Update();

            if (hardware.HardwareType == HardwareType.Cpu)
            {
                foreach (ISensor s in hardware.Sensors)
                {
                    if (s.SensorType == SensorType.Temperature && s.Name.Contains("Core"))
                        cpuTemp = s.Value;
                    if (s.SensorType == SensorType.Clock && s.Name == "CPU Core #1")
                        cpuClock = s.Value;
                }
            }

            if (hardware.HardwareType == HardwareType.GpuNvidia || hardware.HardwareType == HardwareType.GpuAmd || hardware.HardwareType == HardwareType.GpuIntel)
            {
                foreach (ISensor s in hardware.Sensors)
                {
                    if (s.SensorType == SensorType.Temperature)
                        gpuTemp = s.Value;
                    if (s.SensorType == SensorType.Load && s.Name == "GPU Core")
                        gpuLoad = s.Value;
                }
            }
        }

        computer.Close();

        //  Battery Status
        string batteryStatus = GetBatteryStatus();

        //  Report
        Console.WriteLine("\n===== Energy Report =====");
        Console.WriteLine($"Process           : {objProcess.ProcessName} (PID {objProcess.Id})");
        Console.WriteLine($"CPU Usage         : {cpuUsage:F2}%");
        Console.WriteLine($"CPU Temp          : {cpuTemp?.ToString("F1") ?? "N/A"} °C");
        Console.WriteLine($"CPU Clock Speed   : {cpuClock?.ToString("F1") ?? "N/A"} MHz");
        Console.WriteLine($"Memory Usage      : {memoryMB:F2} MB");
        Console.WriteLine($"Disk I/O          : {diskKBs:F2} KB/sec");
        Console.WriteLine($"Disk Queue Len    : {diskQueueLen:F2}");
        Console.WriteLine($"GPU Temp          : {gpuTemp?.ToString("F1") ?? "N/A"} °C");
        Console.WriteLine($"GPU Load          : {gpuLoad?.ToString("F1") ?? "N/A"} %");
        Console.WriteLine($"Battery Status    : {batteryStatus}");
        Console.WriteLine($"Energy (kWh)      : {energyKWh:F6}");
        Console.WriteLine($"CO₂ Emissions     : {co2Grams:F2} g");
        //Console.WriteLine($"Renewable Energy  : {renewablePercent}%");
        Console.WriteLine("============================");

        //  try { process.Kill(); } catch { }
        #endregion

        Console.Read();
    }

    static string GetBatteryStatus()
    {
        try
        {
            var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Battery");
            foreach (var obj in searcher.Get())
            {
                return $"Charge: {obj["EstimatedChargeRemaining"]}% | Status: {obj["BatteryStatus"]}";
            }
        }
        catch { }
        return "Battery info not available";
    }

    static Process GetOrStartProcessById(int processId)
    {
        var runningProcess = Process.GetProcessById(processId);
        try
        {
            return runningProcess;
        }
        catch (Exception ex)
        {
            AnsiConsole.WriteException(ex, new ExceptionSettings
            {
                Format = ExceptionFormats.ShortenEverything | ExceptionFormats.ShowLinks,
                Style = new ExceptionStyle
                {
                    Exception = new Style().Foreground(Color.Grey),
                    Message = new Style().Foreground(Color.White),
                    NonEmphasized = new Style().Foreground(Color.Cornsilk1),
                    Parenthesis = new Style().Foreground(Color.Cornsilk1),
                    Method = new Style().Foreground(Color.Red),
                    ParameterName = new Style().Foreground(Color.Cornsilk1),
                    ParameterType = new Style().Foreground(Color.Red),
                    Path = new Style().Foreground(Color.Red),
                    LineNumber = new Style().Foreground(Color.Cornsilk1),
                }

            });
            return null;
        }
    }

    static int ExtractProcessId(string strProcessName)
    {
        int startIndex = strProcessName.IndexOf(':') + 1;
        int endIndex = strProcessName.IndexOf(')', startIndex);
        return Convert.ToInt32(strProcessName.Substring(startIndex, endIndex - startIndex));
    }

    static string GetProcessInstanceName(int pid)
    {
        var category = new PerformanceCounterCategory("Process");
        foreach (string name in category.GetInstanceNames())
        {
            try
            {
                using var cnt = new PerformanceCounter("Process", "ID Process", name, true);
                if ((int)cnt.RawValue == pid) return name;
            }
            catch { }
        }
        throw new Exception("Instance name not found.");
    }
}