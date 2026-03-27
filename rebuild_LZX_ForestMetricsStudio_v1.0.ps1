# ======================================================================
# 林智创新 · ForestMetrics Studio  (Software v1.0.0)
# Compute Kernel: ForestMetrics V5
# FULL REBUILD + PUBLISH (net8.0, Windows)
#
# Root:   E:\ForestMetricsV5
# Proj:   E:\ForestMetricsV5\ForestMetricsV5
# Output: E:\ForestMetricsV5\_release\林智创新_ForestMetricsStudio_v1.0.exe
#
# Patch v1.0.0-fix1:
#   - Fix ONLY UI numeric "floating tail" display (DBH 0.116999999999...)
#   - Computation & export remain unchanged.
# ======================================================================

Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
$ErrorActionPreference = "Stop"

# ---------------- Config ----------------
$ROOT = "E:\ForestMetricsV5"
$APP  = "ForestMetricsV5"                       # project folder & assembly base
$TEAM = "林智创新"
$PRODUCT = "ForestMetrics Studio"
$SW_VER = "1.0.0"
$KERNEL_VER = "V5"
$EXE_NAME = "${TEAM}_ForestMetricsStudio_v1.0.exe"

$projDir = Join-Path $ROOT $APP
$releaseDir = Join-Path $ROOT "_release"

Write-Host "== $TEAM · $PRODUCT | software v$SW_VER | kernel $KERNEL_VER ==" -ForegroundColor Cyan

# ---------------- Prep ----------------
New-Item -ItemType Directory -Force $ROOT | Out-Null
New-Item -ItemType Directory -Force $releaseDir | Out-Null
Set-Location $ROOT

# ---------------- Ensure templates ----------------
Write-Host "[0/10] Install/Update Avalonia templates..." -ForegroundColor Cyan
dotnet new install Avalonia.Templates --force | Out-Null

# ---------------- Backup old project if exists ----------------
if (Test-Path $projDir) {
  $bak = "${projDir}_bak_{0}" -f (Get-Date -Format "yyyyMMdd_HHmmss")
  Write-Host "[Backup] $projDir -> $bak" -ForegroundColor Yellow
  Move-Item -Force $projDir $bak
}

# ---------------- Create project (no restore) ----------------
Write-Host "[1/10] dotnet new avalonia.app (no restore)..." -ForegroundColor Cyan
dotnet new avalonia.app -n $APP --no-restore | Out-Null
Set-Location $projDir

# ---------------- Patch csproj to net8.0 + version metadata ----------------
Write-Host "[2/10] Patch csproj to net8.0 + version metadata..." -ForegroundColor Cyan
$csproj = Get-ChildItem -Filter "*.csproj" | Select-Object -First 1
if (-not $csproj) { throw "找不到 csproj。" }

$raw = Get-Content $csproj.FullName -Raw

# TargetFramework(s) -> net8.0
$raw = $raw -replace "<TargetFrameworks>.*?</TargetFrameworks>", "<TargetFramework>net8.0</TargetFramework>"
$raw = $raw -replace "<TargetFramework>.*?</TargetFramework>", "<TargetFramework>net8.0</TargetFramework>"

# LangVersion -> latest
if ($raw -match "<LangVersion>preview</LangVersion>") {
  $raw = $raw -replace "<LangVersion>preview</LangVersion>", "<LangVersion>latest</LangVersion>"
} elseif ($raw -notmatch "<LangVersion>") {
  $raw = $raw -replace "</PropertyGroup>", "  <LangVersion>latest</LangVersion>`r`n</PropertyGroup>"
}

# Add version/company/product metadata
if ($raw -notmatch "<Version>") {
  $raw = $raw -replace "</PropertyGroup>", "  <Version>$SW_VER</Version>`r`n  <AssemblyVersion>$SW_VER</AssemblyVersion>`r`n  <FileVersion>$SW_VER</FileVersion>`r`n  <Company>$TEAM</Company>`r`n  <Product>$TEAM $PRODUCT</Product>`r`n</PropertyGroup>"
} else {
  $raw = $raw -replace "<Version>.*?</Version>", "<Version>$SW_VER</Version>"
  if ($raw -match "<FileVersion>") { $raw = $raw -replace "<FileVersion>.*?</FileVersion>", "<FileVersion>$SW_VER</FileVersion>" }
  if ($raw -match "<AssemblyVersion>") { $raw = $raw -replace "<AssemblyVersion>.*?</AssemblyVersion>", "<AssemblyVersion>$SW_VER</AssemblyVersion>" }
}

Set-Content -Encoding utf8 $csproj.FullName $raw

# ---------------- Add NuGet packages (no restore) ----------------
Write-Host "[3/10] Add NuGet packages..." -ForegroundColor Cyan
dotnet add package Avalonia.Controls.DataGrid --version 11.3.12 --no-restore | Out-Null
dotnet add package OxyPlot.Avalonia --version 2.1.0-Avalonia11 --no-restore | Out-Null
dotnet add package OxyPlot.SkiaSharp --version 2.1.0 --no-restore | Out-Null
dotnet add package ExcelDataReader --version 3.7.0 --no-restore | Out-Null
dotnet add package ExcelDataReader.DataSet --version 3.7.0 --no-restore | Out-Null
dotnet add package System.Text.Encoding.CodePages --version 8.0.0 --no-restore | Out-Null

# ---------------- Create folders ----------------
Write-Host "[4/10] Create folders..." -ForegroundColor Cyan
New-Item -ItemType Directory -Force ".\Services" | Out-Null

# ======================================================================
# [5/10] WRITE ALL SOURCE FILES (NO OMISSION)
# ======================================================================
Write-Host "[5/10] Write ALL source files..." -ForegroundColor Cyan

# ---------------- App.axaml ----------------
@'
<Application xmlns="https://github.com/avaloniaui"
             xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
             x:Class="ForestMetricsV5.App">
  <Application.Styles>
    <FluentTheme />
    <StyleInclude Source="avares://Avalonia.Controls.DataGrid/Themes/Fluent.xaml"/>
    <StyleInclude Source="avares://OxyPlot.Avalonia/Themes/Default.axaml"/>
  </Application.Styles>
</Application>
'@ | Set-Content -Encoding utf8 .\App.axaml

# ---------------- App.axaml.cs ----------------
@'
using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Markup.Xaml;
using Avalonia.Styling;

namespace ForestMetricsV5;

public partial class App : Application
{
    public override void Initialize()
    {
        AvaloniaXamlLoader.Load(this);
        RequestedThemeVariant = ThemeVariant.Light;
    }

    public override void OnFrameworkInitializationCompleted()
    {
        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
            desktop.MainWindow = new MainWindow();

        base.OnFrameworkInitializationCompleted();
    }
}
'@ | Set-Content -Encoding utf8 .\App.axaml.cs

# ---------------- MainWindow.axaml (你的最终定稿原文) ----------------
@'
<Window xmlns="https://github.com/avaloniaui"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:oxy="clr-namespace:OxyPlot.Avalonia;assembly=OxyPlot.Avalonia"
        mc:Ignorable="d"
        x:Class="ForestMetricsV5.MainWindow"
        Title="林智创新 · ForestMetrics Studio"
        Width="1400" Height="820"
        WindowStartupLocation="CenterScreen">


  <Window.Styles>
    <Style Selector="TabItem">
      <Setter Property="FontSize" Value="15"/>
    </Style>
  </Window.Styles>
<Grid RowDefinitions="Auto,*" Margin="14">
    <!-- 顶部标题 -->
    <Border CornerRadius="16" Padding="14"
            BorderBrush="#D9DDE3" BorderThickness="1"
            Background="White">
      <StackPanel>
        <TextBlock Text="林智创新 · ForestMetrics Studio" FontSize="20" FontWeight="SemiBold"/>
        <TextBlock Text="W/M, CI, Openness, U, UCI_D/UCI_H, C, F + Entropy Weights + Q(g) | CORE/BUFFER bbox"
                   Foreground="#666" FontSize="12"/>
      </StackPanel>
    </Border>

    <Grid Grid.Row="1" Margin="0,12,0,0" ColumnDefinitions="440,*" ColumnSpacing="12">
      <!-- 左侧控制面板 -->
      <Border Grid.Column="0" CornerRadius="16" Padding="14"
              BorderBrush="#D9DDE3" BorderThickness="1"
              Background="White">
        <ScrollViewer>
          <StackPanel Spacing="10">

            <TextBlock Text="输入 / 参数 / 导出" FontSize="18" FontWeight="SemiBold"/>

            <StackPanel Spacing="6">
              <TextBlock Text="输入文件（Excel/CSV/TXT）"/>
              <Grid ColumnDefinitions="*,Auto" ColumnSpacing="8">
                <TextBox x:Name="TbInput" IsReadOnly="True"/>
                <Button x:Name="BtnPickInput" Grid.Column="1" Content="选择" Padding="14,6"/>
              </Grid>
            </StackPanel>

            <StackPanel Spacing="6">
              <TextBlock Text="输出文件夹（空=与输入同目录）"/>
              <Grid ColumnDefinitions="*,Auto" ColumnSpacing="8">
                <TextBox x:Name="TbOutDir"/>
                <Button x:Name="BtnPickOutDir" Grid.Column="1" Content="选择" Padding="14,6"/>
              </Grid>
            </StackPanel>

            <Separator/>

            <StackPanel Spacing="6">
              <TextBlock Text="BufferSize（米，0=不启用；bbox 模式）"/>
              <TextBox x:Name="TbBuffer" Text="2"/>
            </StackPanel>

            <StackPanel Spacing="6">
              <TextBlock Text="F 影响圈系数：infl_k × CrownDiameter（默认 1.5）"/>
              <TextBox x:Name="TbInflK" Text="1.5"/>
            </StackPanel>

            <Separator/>

            <TextBlock Text="计算模块（可勾选）" FontSize="16" FontWeight="SemiBold"/>

            <CheckBox x:Name="CkWM" IsChecked="True" Content="W / M"/>
            <CheckBox x:Name="CkLayer" IsChecked="True" Content="层级 / CI / Openness(÷4) / PathOrder"/>
            <CheckBox x:Name="CkUUCI" IsChecked="True" Content="U + UCI(主=UCI_D) / UCI_H / UCI_class"/>
            <CheckBox x:Name="CkCF" IsChecked="True" Content="C + F（CrownRadius=冠径直径）"/>
            <CheckBox x:Name="CkQ" IsChecked="True" Content="熵权 + Q(g)（CORE 估权；ALL/CORE/BUFFER 均值）"/>

            <Separator/>

            <TextBlock Text="导出（可选）" FontSize="16" FontWeight="SemiBold"/>

            <WrapPanel>
              <CheckBox x:Name="CkExpMain" IsChecked="True" Content="主表" Margin="0,0,14,6"/>
              <CheckBox x:Name="CkExpMeans" IsChecked="True" Content="林分均值" Margin="0,0,14,6"/>
              <CheckBox x:Name="CkExpEntropy" IsChecked="True" Content="熵权表" Margin="0,0,14,6"/>
              <CheckBox x:Name="CkExpQTable" IsChecked="True" Content="综合指数表" Margin="0,0,14,6"/>
            </WrapPanel>

            <StackPanel Spacing="6">
              <TextBlock Text="导出格式"/>
              <ComboBox x:Name="CbFormat" SelectedIndex="0" Width="200">
                <ComboBoxItem Content="TXT（制表符）"/>
                <ComboBoxItem Content="CSV（逗号）"/>
              </ComboBox>
            </StackPanel>

            <Separator/>

            <Grid ColumnDefinitions="*,*,*" ColumnSpacing="8">
              <Button x:Name="BtnRun" Content="开始计算" Height="40"/>
              <Button x:Name="BtnEnlarge" Grid.Column="1" Content="放大图表" Height="40"/>
              <Button x:Name="BtnExportPng" Grid.Column="2" Content="导出图表PNG" Height="40"/>
            </Grid>

            <ProgressBar x:Name="Pb" Minimum="0" Maximum="100" Height="10" Value="0"/>

            <TextBlock Text="日志" FontSize="14" FontWeight="SemiBold"/>
            <TextBox x:Name="TbLog" AcceptsReturn="True" IsReadOnly="True" Height="240"
                     TextWrapping="Wrap" ScrollViewer.VerticalScrollBarVisibility="Auto"/>
          </StackPanel>
        </ScrollViewer>
      </Border>

      <!-- 右侧 Tab 区 -->
      <Border Grid.Column="1" CornerRadius="16" Padding="12"
              BorderBrush="#D9DDE3" BorderThickness="1"
              Background="White">
        <TabControl x:Name="Tabs" FontSize="10">
          <TabItem Header="Preview">
            <DataGrid x:Name="GridPreview" IsReadOnly="True" GridLinesVisibility="All" AutoGenerateColumns="False"/>
          </TabItem>

          <TabItem Header="Voronoi">
            <oxy:PlotView x:Name="PlotVoronoi"/>
          </TabItem>

          <TabItem Header="Distributions">
            <Grid RowDefinitions="Auto,*">
              <StackPanel Orientation="Horizontal" Spacing="10" Margin="0,0,0,10">
                <TextBlock Text="指标：" VerticalAlignment="Center"/>
                <ComboBox x:Name="CbMetric" Width="220"/>
                <TextBlock Text="分组：" VerticalAlignment="Center"/>
                <ComboBox x:Name="CbGroup" Width="140">
                  <ComboBoxItem Content="ALL"/>
                  <ComboBoxItem Content="CORE"/>
                  <ComboBoxItem Content="BUFFER"/>
                </ComboBox>
                <Button x:Name="BtnRefreshDist" Content="刷新"/>
              </StackPanel>

              <oxy:PlotView x:Name="PlotDist" Grid.Row="1"/>
            </Grid>
          </TabItem>

          <TabItem Header="Stand Means">
            <DataGrid x:Name="GridMeans" IsReadOnly="True" GridLinesVisibility="All" AutoGenerateColumns="False"/>
          </TabItem>

          <TabItem Header="Entropy Weights">
            <DataGrid x:Name="GridEntropy" IsReadOnly="True" GridLinesVisibility="All" AutoGenerateColumns="False"/>
          </TabItem>

          <TabItem Header="Q Table">
            <DataGrid x:Name="GridQ" IsReadOnly="True" GridLinesVisibility="All" AutoGenerateColumns="False"/>
          </TabItem>
        </TabControl>
      </Border>
    </Grid>
  </Grid>
</Window>
'@ | Set-Content -Encoding utf8 .\MainWindow.axaml

# ---------------- MainWindow.axaml.cs ----------------
# 重要：不要 using Avalonia.Controls.DataGrid; （会触发 CS0138）
# Patch ONLY: BindGrid uses FormatCell() to avoid floating tails.
@'
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Data;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
using OxyPlot;
using ForestMetricsV5.Services;

namespace ForestMetricsV5;

public partial class MainWindow : Window
{
    private string _inputPath = "";
    private DataTable? _computed;
    private DataTable? _previewMain;
    private DataTable? _means;
    private DataTable? _entropy;
    private DataTable? _qtable;

    public MainWindow()
    {
        InitializeComponent();

        BtnPickInput.Click += async (_, __) => await PickInputAsync();
        BtnPickOutDir.Click += async (_, __) => await PickOutDirAsync();
        BtnRun.Click += async (_, __) => await RunAsync();

        BtnRefreshDist.Click += (_, __) => RefreshDistributionPlot();
        CbMetric.SelectionChanged += (_, __) => RefreshDistributionPlot();
        CbGroup.SelectionChanged += (_, __) => RefreshDistributionPlot();

        BtnEnlarge.Click += (_, __) => EnlargeActivePlot();
        BtnExportPng.Click += async (_, __) => await ExportActivePlotPngAsync();

        AppendLog("[UI] Ready.");
    }

    private void AppendLog(string s)
    {
        Dispatcher.UIThread.Post(() =>
        {
            TbLog.Text = (TbLog.Text ?? "") + s + Environment.NewLine;
            TbLog.CaretIndex = TbLog.Text?.Length ?? 0;
        });
    }

    private async Task PickInputAsync()
    {
        var types = new[]
        {
            new FilePickerFileType("Data") { Patterns = new[] { "*.xlsx","*.xls","*.xlsm","*.xlsb","*.csv","*.txt" } },
            FilePickerFileTypes.All
        };

        var files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "选择数据文件（Excel/CSV/TXT）",
            AllowMultiple = false,
            FileTypeFilter = types
        });

        var f = files.FirstOrDefault();
        var path = f?.TryGetLocalPath();
        if (string.IsNullOrWhiteSpace(path)) return;

        _inputPath = path;
        TbInput.Text = _inputPath;
        AppendLog($"[Input] {_inputPath}");

        try
        {
            var raw = DataIO.ReadTableFlexible(_inputPath);
            var std = ColumnMapper.MapToStandard(raw);
            BindGrid(GridPreview, std, 200);
            AppendLog($"[Preview] 映射成功：rows={std.Rows.Count}（显示前200行）");
        }
        catch (Exception ex)
        {
            AppendLog("[Preview] FAILED: " + ex.Message);
        }
    }

    private async Task PickOutDirAsync()
    {
        var folders = await StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions
        {
            Title = "选择输出文件夹",
            AllowMultiple = false
        });

        var f = folders.FirstOrDefault();
        var path = f?.TryGetLocalPath();
        if (string.IsNullOrWhiteSpace(path)) return;

        TbOutDir.Text = path;
        AppendLog($"[OutDir] {path}");
    }

    private async Task RunAsync()
    {
        if (string.IsNullOrWhiteSpace(_inputPath) || !File.Exists(_inputPath))
        {
            AppendLog("[Run] 请先选择有效输入文件。");
            return;
        }

        BtnRun.IsEnabled = false;
        Pb.Value = 0;

        bool useWM     = CkWM.IsChecked ?? true;
        bool useLayer  = CkLayer.IsChecked ?? true;
        bool useUUCI   = CkUUCI.IsChecked ?? true;
        bool useCF     = CkCF.IsChecked ?? true;
        bool useQ      = CkQ.IsChecked ?? true;

        bool expMain   = CkExpMain.IsChecked ?? true;
        bool expMeans  = CkExpMeans.IsChecked ?? true;
        bool expEnt    = CkExpEntropy.IsChecked ?? true;
        bool expQTab   = CkExpQTable.IsChecked ?? true;

        var outDir = (TbOutDir.Text ?? "").Trim();
        var format = (CbFormat.SelectedIndex == 1) ? "CSV" : "TXT";

        double bufferSize = ParseDouble(TbBuffer.Text, 2.0);
        double inflK      = ParseDouble(TbInflK.Text, 1.5);

        AppendLog("---- RUN START ----");

        try
        {
            var res = await Task.Run(() =>
            {
                UIProg(10);
                var raw = DataIO.ReadTableFlexible(_inputPath);

                UIProg(25);
                var std = ColumnMapper.MapToStandard(raw);

                UIProg(35);
                BufferZone.AddBBoxZoneColumns(std, bufferSize);

                UIProg(60);
                var computed = MetricsEngine.ComputeAll(std, inflK);

                UIProg(70);
                DataTable ew;
                DataTable qtable;
                double qAll = double.NaN, qCore = double.NaN, qBuf = double.NaN;

                if (useQ)
                {
                    (ew, qtable, qAll, qCore, qBuf) = EntropyQ.ComputeEntropyAndQ(computed);
                }
                else
                {
                    ew = new DataTable("EntropyWeights");
                    ew.Columns.Add("Metric", typeof(string));
                    ew.Columns.Add("Weight", typeof(double));

                    qtable = new DataTable("QTable");
                    qtable.Columns.Add("Tree", typeof(double));
                    qtable.Columns.Add("Q", typeof(double));
                }

                UIProg(78);
                var means = StandMeans.BuildStandMeans(computed, qAll, qCore, qBuf);

                UIProg(85);
                var mainOut = ColumnFilter.BuildMainForPreview(computed, useWM, useLayer, useUUCI, useCF, useQ);

                UIProg(90);
                var paths = Exporter.WriteAll(
                    mainOut, means, ew, qtable,
                    _inputPath, outDir, format,
                    expMain, expMeans, expEnt, expQTab
                );

                UIProg(94);
                var metricList = MetricCatalog.BuildMetricList(useWM, useLayer, useUUCI, useCF, useQ);

                return (computed, mainOut, means, ew, qtable, paths, metricList);
            });

            _computed    = res.computed;
            _previewMain = res.mainOut;
            _means       = res.means;
            _entropy     = res.ew;
            _qtable      = res.qtable;

            BindGrid(GridPreview, _previewMain, 600);
            BindGrid(GridMeans, _means, 600);
            BindGrid(GridEntropy, _entropy, 200);
            BindGrid(GridQ, _qtable, 800);

            PlotVoronoi.Model = VoronoiPlotter.BuildVoronoiApproxModel(
                _computed,
                title: $"Voronoi (approx) | Buffer={ParseDouble(TbBuffer.Text, 2.0)}m"
            );
            PlotVoronoi.InvalidatePlot(true);

            CbMetric.ItemsSource = res.metricList;
            if (res.metricList.Count > 0 && CbMetric.SelectedIndex < 0) CbMetric.SelectedIndex = 0;
            if (CbGroup.SelectedIndex < 0) CbGroup.SelectedIndex = 0;
            RefreshDistributionPlot();

            if (res.paths.Main != null) AppendLog($"[Export] {res.paths.Main}");
            if (res.paths.Means != null) AppendLog($"[Export] {res.paths.Means}");
            if (res.paths.Entropy != null) AppendLog($"[Export] {res.paths.Entropy}");
            if (res.paths.QTree != null) AppendLog($"[Export] {res.paths.QTree}");

            Pb.Value = 100;
            AppendLog("---- RUN DONE ----");
        }
        catch (Exception ex)
        {
            AppendLog("[Run] FAILED: " + ex);
        }
        finally
        {
            BtnRun.IsEnabled = true;
        }

        void UIProg(double v) => Dispatcher.UIThread.Post(() => Pb.Value = v);
    }

    private void RefreshDistributionPlot()
    {
        if (_computed == null) return;
        if (CbMetric.SelectedItem is not string metric) return;

        string group = "ALL";
        if (CbGroup.SelectedItem is ComboBoxItem cbi)
            group = cbi.Content?.ToString() ?? "ALL";
        else if (CbGroup.SelectedItem is string s)
            group = s;

        PlotDist.Model = HistPlotter.BuildMetricHistogram(_computed, metric, group);
        PlotDist.InvalidatePlot(true);
    }

    private void EnlargeActivePlot()
    {
        var tab = Tabs.SelectedItem as TabItem;
        var header = tab?.Header?.ToString() ?? "";

        PlotModel? model =
            header == "Voronoi" ? PlotVoronoi.Model :
            header == "Distributions" ? PlotDist.Model :
            null;

        if (model == null)
        {
            AppendLog("[Plot] 当前Tab没有可放大的图。");
            return;
        }

        var win = new Window
        {
            Title = header,
            Width = 1200,
            Height = 860,
            WindowStartupLocation = WindowStartupLocation.CenterScreen,
            Content = new OxyPlot.Avalonia.PlotView { Model = model }
        };
        win.Show();
    }

    private async Task ExportActivePlotPngAsync()
    {
        var tab = Tabs.SelectedItem as TabItem;
        var header = tab?.Header?.ToString() ?? "";

        PlotModel? model =
            header == "Voronoi" ? PlotVoronoi.Model :
            header == "Distributions" ? PlotDist.Model :
            null;

        if (model == null)
        {
            AppendLog("[PNG] 当前Tab没有可导出的图。");
            return;
        }

        var file = await StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
        {
            Title = "导出 PNG",
            SuggestedFileName = $"{header}_{DateTime.Now:yyyyMMdd_HHmmss}.png",
            DefaultExtension = "png"
        });

        var path = file?.TryGetLocalPath();
        if (string.IsNullOrWhiteSpace(path)) return;

        bool ok = PlotExport.TryExportPng(model, path, 1800, 1200);
        AppendLog(ok ? $"[PNG] {path}" : "[PNG] 导出失败（检查 OxyPlot.SkiaSharp 包）。");
    }

    // ---------------- PATCH ONLY (UI numeric formatting) ----------------
    private static string FormatCell(object? v)
    {
        if (v is null || v is DBNull) return "";

        if (v is bool b) return b ? "True" : "False";

        if (v is double d)
        {
            if (!double.IsFinite(d)) return "";
            return d.ToString("0.######", CultureInfo.InvariantCulture); // avoid 0.116999999999 tails
        }
        if (v is float f)
        {
            if (!float.IsFinite(f)) return "";
            return f.ToString("0.######", CultureInfo.InvariantCulture);
        }
        if (v is decimal m)
        {
            return m.ToString("0.######", CultureInfo.InvariantCulture);
        }

        return v.ToString() ?? "";
    }

    private void BindGrid(Avalonia.Controls.DataGrid grid, DataTable table, int maxRows)
    {
        Dispatcher.UIThread.Post(() =>
        {
            grid.Columns.Clear();

            int n = Math.Min(table.Rows.Count, maxRows);
            var cols = table.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToList();

            var items = new List<Dictionary<string, string>>(n);
            for (int i = 0; i < n; i++)
            {
                var r = table.Rows[i];
                var d = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var c in cols)
                {
                    var v = r[c];
                    d[c] = FormatCell(v); // PATCH: format numeric for UI
                }
                items.Add(d);
            }

            foreach (var c in cols)
            {
                grid.Columns.Add(new Avalonia.Controls.DataGridTextColumn
                {
                    Header = c,
                    Binding = new Binding($"[{c}]")
                });
            }

            grid.ItemsSource = items;
        });
    }
    // ------------------------------------------------------------------

    private static double ParseDouble(string? s, double fallback)
    {
        s = (s ?? "").Trim();
        if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var x)) return x;
        if (double.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out x)) return x;
        return fallback;
    }
}
'@ | Set-Content -Encoding utf8 .\MainWindow.axaml.cs

# ---------------- Services: DataIO.cs ----------------
@'
using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ExcelDataReader;

namespace ForestMetricsV5.Services;

public static class DataIO
{
    public static DataTable ReadTableFlexible(string path)
    {
        if (!File.Exists(path)) throw new FileNotFoundException(path);

        var ext = Path.GetExtension(path).ToLowerInvariant();
        bool isExcel = ext is ".xlsx" or ".xls" or ".xlsm" or ".xlsb";

        return isExcel ? ReadExcel(path) : ReadText(path);
    }

    private static DataTable ReadExcel(string path)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        DataTable dt = ReadExcelCore(path, useHeader: true);

        bool isNumHeader = dt.Columns.Cast<DataColumn>()
            .All(c => c.ColumnName.Length > 0 && c.ColumnName.All(char.IsDigit));

        if (isNumHeader)
        {
            dt = ReadExcelCore(path, useHeader: false);
            for (int i = 0; i < dt.Columns.Count; i++)
                dt.Columns[i].ColumnName = $"Var{i + 1}";
        }

        for (int i = 0; i < dt.Columns.Count; i++)
        {
            if (string.IsNullOrWhiteSpace(dt.Columns[i].ColumnName))
                dt.Columns[i].ColumnName = $"Var{i + 1}";
        }

        return dt;
    }

    private static DataTable ReadExcelCore(string path, bool useHeader)
    {
        using var fs = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var reader = ExcelReaderFactory.CreateReader(fs);

        var conf = new ExcelDataSetConfiguration
        {
            ConfigureDataTable = _ => new ExcelDataTableConfiguration
            {
                UseHeaderRow = useHeader
            }
        };

        var ds = reader.AsDataSet(conf);
        if (ds.Tables.Count == 0) throw new Exception("Excel 无可用工作表。");
        return ds.Tables[0];
    }

    private static DataTable ReadText(string path)
    {
        var lines = File.ReadAllLines(path, Encoding.UTF8)
            .Where(s => !string.IsNullOrWhiteSpace(s))
            .ToArray();
        if (lines.Length == 0) throw new Exception("文本为空。");

        var first = lines[0].Trim();
        char? delim = DetectDelimiter(first);

        string[] SplitLine(string s)
        {
            s = s.Trim();
            if (delim == null) return Regex.Split(s, @"\s+").Where(x => x.Length > 0).ToArray();
            return s.Split(delim.Value).Select(x => x.Trim()).ToArray();
        }

        var firstTokens = SplitLine(first);
        bool headerLike = firstTokens.Any(tok => tok.Any(ch => char.IsLetter(ch) || (ch >= 0x4E00 && ch <= 0x9FFF)));

        var dt = new DataTable("Text");
        int cols = firstTokens.Length;

        if (headerLike)
        {
            for (int i = 0; i < cols; i++)
            {
                var name = string.IsNullOrWhiteSpace(firstTokens[i]) ? $"Var{i + 1}" : firstTokens[i];
                dt.Columns.Add(MakeUnique(dt, name), typeof(string));
            }

            for (int r = 1; r < lines.Length; r++)
            {
                var tok = SplitLine(lines[r]);
                if (tok.Length < cols) tok = tok.Concat(Enumerable.Repeat("", cols - tok.Length)).ToArray();
                var row = dt.NewRow();
                for (int c = 0; c < cols; c++) row[c] = tok[c];
                dt.Rows.Add(row);
            }
        }
        else
        {
            for (int i = 0; i < cols; i++) dt.Columns.Add($"Var{i + 1}", typeof(string));

            for (int r = 0; r < lines.Length; r++)
            {
                var tok = SplitLine(lines[r]);
                if (tok.Length < cols) tok = tok.Concat(Enumerable.Repeat("", cols - tok.Length)).ToArray();
                var row = dt.NewRow();
                for (int c = 0; c < cols; c++) row[c] = tok[c];
                dt.Rows.Add(row);
            }
        }

        return dt;
    }

    private static char? DetectDelimiter(string line)
    {
        if (line.Contains('\t')) return '\t';
        if (line.Contains(',')) return ',';
        if (line.Contains(';')) return ';';
        return null;
    }

    private static string MakeUnique(DataTable dt, string name)
    {
        var n = name;
        int k = 1;
        while (dt.Columns.Contains(n))
        {
            n = $"{name}_{k++}";
        }
        return n;
    }
}
'@ | Set-Content -Encoding utf8 .\Services\DataIO.cs

# ---------------- Services: ColumnMapper.cs ----------------
@'
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace ForestMetricsV5.Services;

public static class ColumnMapper
{
    public static DataTable MapToStandard(DataTable raw)
    {
        var normNames = raw.Columns.Cast<DataColumn>()
            .Select(c => Normalize(c.ColumnName))
            .ToArray();

        var aliases = new Dictionary<string, string[]>
        {
            ["tree"]    = new[] {"tree","tag","id","树号","编号","treeid","no"},
            ["species"] = new[] {"species","sp","树种","speciesname","spname","种"},
            ["dbh"]     = new[] {"dbh","d","胸径","径","dbhcm","d1.3","d13","d130"},
            ["height"]  = new[] {"height","h","树高","高","ht"},
            ["crownd"]  = new[] {"crowndiameter","crdiameter","cd","冠径","冠幅","crownd","crownwidth"},
            ["crownr"]  = new[] {"crownradius","cr","冠半径","冠幅半径","crownr"},
            ["x"]       = new[] {"x","coordx","lon","xcoord","east","coordinatex","横坐标"},
            ["y"]       = new[] {"y","coordy","lat","ycoord","north","coordinatey","纵坐标"}
        };

        int Find(string key)
        {
            foreach (var a in aliases[key])
                for (int i = 0; i < normNames.Length; i++)
                    if (normNames[i] == a) return i;
            return -1;
        }

        int iTree = Find("tree");
        int iX = Find("x");
        int iY = Find("y");
        int iDBH = Find("dbh");
        int iH = Find("height");
        int iSp = Find("species");
        int iCR = Find("crownr");
        int iCD = Find("crownd");

        bool hasA = (iTree >= 0 && iX >= 0 && iY >= 0 && iDBH >= 0 && iH >= 0);

        var dt = NewStandardTable();
        var treeEnc = new CategoryEncoder(1);
        var spEnc = new CategoryEncoder(1);

        if (hasA)
        {
            for (int r = 0; r < raw.Rows.Count; r++)
            {
                var row = raw.Rows[r];

                double tree = ToDouble(row[iTree]);
                if (double.IsNaN(tree)) tree = treeEnc.Encode(row[iTree]?.ToString() ?? "");

                double x = ToDouble(row[iX]);
                double y = ToDouble(row[iY]);
                double dbh = ToDouble(row[iDBH]);
                double h = ToDouble(row[iH]);
                if (double.IsNaN(x) || double.IsNaN(y) || double.IsNaN(dbh) || double.IsNaN(h)) continue;

                double sp = (iSp >= 0) ? ToDoubleOrCategory(row[iSp], spEnc) : 0;

                double crownDia;
                if (iCR >= 0) crownDia = ToDouble(row[iCR]);
                else if (iCD >= 0) crownDia = ToDouble(row[iCD]);
                else crownDia = 0.3 * h;

                dt.Rows.Add(tree, x, y, dbh, sp, h, crownDia);
            }

            PostProcessDbhCmToM(dt);
            return dt;
        }

        if (raw.Columns.Count >= 8)
        {
            for (int r = 0; r < raw.Rows.Count; r++)
            {
                var row = raw.Rows[r];

                double tree = ToDouble(row[1]);
                if (double.IsNaN(tree)) tree = treeEnc.Encode(row[1]?.ToString() ?? "");

                double sp = ToDoubleOrCategory(row[2], spEnc);
                double dbh = ToDouble(row[3]);
                double h = ToDouble(row[4]);
                double crownDia = ToDouble(row[5]);
                double x = ToDouble(row[6]);
                double y = ToDouble(row[7]);
                if (double.IsNaN(x) || double.IsNaN(y) || double.IsNaN(dbh) || double.IsNaN(h)) continue;

                dt.Rows.Add(tree, x, y, dbh, sp, h, crownDia);
            }

            PostProcessDbhCmToM(dt);
            return dt;
        }

        throw new Exception("无法识别列名/列顺序。支持: A) Tree/X/Y/DBH/Species/Height/[CrownRadius|CrownDiameter] 或 B) Plot/Tag/SP/D/H/CR/X/Y");
    }

    private static DataTable NewStandardTable()
    {
        var dt = new DataTable("Standard");
        dt.Columns.Add("Tree", typeof(double));
        dt.Columns.Add("X", typeof(double));
        dt.Columns.Add("Y", typeof(double));
        dt.Columns.Add("DBH", typeof(double));
        dt.Columns.Add("Species", typeof(double));
        dt.Columns.Add("Height", typeof(double));
        dt.Columns.Add("CrownRadius", typeof(double)); // 实际装“冠径直径”
        return dt;
    }

    private static void PostProcessDbhCmToM(DataTable dt)
    {
        var dbhs = dt.AsEnumerable().Select(r => (double)r["DBH"]).Where(v => !double.IsNaN(v)).ToArray();
        if (dbhs.Length == 0) return;
        Array.Sort(dbhs);
        double med = dbhs[dbhs.Length / 2];
        if (med > 2 && med < 200)
        {
            foreach (DataRow r in dt.Rows)
                r["DBH"] = (double)r["DBH"] / 100.0;
        }
    }

    private static string Normalize(string s)
    {
        s = (s ?? "").Trim().ToLowerInvariant();
        return Regex.Replace(s, @"[^a-z0-9\u4e00-\u9fff]+", "");
    }

    private static double ToDouble(object? v)
    {
        if (v is null || v is DBNull) return double.NaN;
        if (v is double d) return d;
        if (v is float f) return f;
        if (v is int i) return i;
        if (v is long l) return l;

        var s = v.ToString()?.Trim();
        if (string.IsNullOrWhiteSpace(s)) return double.NaN;

        if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var x)) return x;
        if (double.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out x)) return x;
        return double.NaN;
    }

    private static double ToDoubleOrCategory(object? v, CategoryEncoder enc)
    {
        var x = ToDouble(v);
        if (!double.IsNaN(x)) return x;
        return enc.Encode(v?.ToString() ?? "");
    }

    private sealed class CategoryEncoder
    {
        private readonly Dictionary<string, int> _map = new(StringComparer.OrdinalIgnoreCase);
        private int _next;
        public CategoryEncoder(int startFrom) { _next = startFrom; }
        public int Encode(string key)
        {
            key = (key ?? "").Trim();
            if (key.Length == 0) return 0;
            if (_map.TryGetValue(key, out var v)) return v;
            v = _next++;
            _map[key] = v;
            return v;
        }
    }
}
'@ | Set-Content -Encoding utf8 .\Services\ColumnMapper.cs

# ---------------- Services: BufferZone.cs ----------------
@'
using System;
using System.Data;
using System.Linq;

namespace ForestMetricsV5.Services;

public static class BufferZone
{
    public static void AddBBoxZoneColumns(DataTable t, double bufferSizeM)
    {
        EnsureCol(t, "BufferSize_m", typeof(double));
        EnsureCol(t, "DistToBoundary", typeof(double));
        EnsureCol(t, "IsCore", typeof(bool));
        EnsureCol(t, "Zone", typeof(string));

        double xmin = t.AsEnumerable().Min(r => (double)r["X"]);
        double xmax = t.AsEnumerable().Max(r => (double)r["X"]);
        double ymin = t.AsEnumerable().Min(r => (double)r["Y"]);
        double ymax = t.AsEnumerable().Max(r => (double)r["Y"]);

        for (int i = 0; i < t.Rows.Count; i++)
        {
            var r = t.Rows[i];
            double x = (double)r["X"];
            double y = (double)r["Y"];
            double dL = x - xmin;
            double dR = xmax - x;
            double dB = y - ymin;
            double dT = ymax - y;
            double dist = Math.Min(Math.Min(dL, dR), Math.Min(dB, dT));

            bool isCore = bufferSizeM <= 0 ? true : dist >= bufferSizeM;

            r["BufferSize_m"] = bufferSizeM;
            r["DistToBoundary"] = dist;
            r["IsCore"] = isCore;
            r["Zone"] = isCore ? "CORE" : "BUFFER";
        }
    }

    private static void EnsureCol(DataTable t, string name, Type type)
    {
        if (!t.Columns.Contains(name)) t.Columns.Add(name, type);
    }
}
'@ | Set-Content -Encoding utf8 .\Services\BufferZone.cs

# ---------------- Services: MetricsEngine.cs ----------------
@'
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace ForestMetricsV5.Services;

public static class MetricsEngine
{
    public static DataTable ComputeAll(DataTable standard, double inflK)
    {
        if (standard.Rows.Count < 5) throw new Exception("至少需要 5 株树（n>=5）才能计算四近邻指标。");

        int n = standard.Rows.Count;

        var X = new double[n];
        var Y = new double[n];
        var Tree = new double[n];
        var DBH = new double[n];   // m
        var Sp = new double[n];
        var H = new double[n];
        var CD = new double[n];    // CrownRadius 字段里装的是“冠径直径”

        for (int i = 0; i < n; i++)
        {
            var r = standard.Rows[i];
            Tree[i] = (double)r["Tree"];
            X[i] = (double)r["X"];
            Y[i] = (double)r["Y"];
            DBH[i] = (double)r["DBH"];
            Sp[i] = (double)r["Species"];
            H[i] = (double)r["Height"];
            CD[i] = (double)r["CrownRadius"];
        }

        var pts = new (double x, double y)[n];
        for (int i = 0; i < n; i++) pts[i] = (X[i], Y[i]);
        var kd = new KDTree2D(pts);

        int[,] nn = new int[n, 4];
        double[,] dd = new double[n, 4];

        for (int i = 0; i < n; i++)
        {
            var res = kd.KNearest(i, 6).Where(p => p.index != i).Take(4).ToArray();
            if (res.Length < 4) throw new Exception("近邻不足（重复点或极端分布）。");

            for (int k = 0; k < 4; k++)
            {
                nn[i, k] = res[k].index;
                dd[i, k] = Math.Sqrt(res[k].dist2);
            }
        }

        var outT = standard.Copy();

        void Add(string name, Type t) { if (!outT.Columns.Contains(name)) outT.Columns.Add(name, t); }

        Add("Tree1", typeof(double)); Add("Tree2", typeof(double)); Add("Tree3", typeof(double)); Add("Tree4", typeof(double));
        Add("Dist1", typeof(double)); Add("Dist2", typeof(double)); Add("Dist3", typeof(double)); Add("Dist4", typeof(double));

        Add("W", typeof(double)); Add("M", typeof(double));

        Add("DominantHeight", typeof(double));
        Add("LayerCategory", typeof(double));
        Add("LayerCategory1", typeof(double));
        Add("LayerCategory2", typeof(double));
        Add("LayerCategory3", typeof(double));
        Add("LayerCategory4", typeof(double));
        Add("LayerCount", typeof(double));
        Add("LayerIndex", typeof(double));
        Add("Discreteness", typeof(double));
        Add("Openness", typeof(double));
        Add("K", typeof(double));
        Add("CI1", typeof(double)); Add("CI2", typeof(double)); Add("CI3", typeof(double)); Add("CI4", typeof(double)); Add("CI", typeof(double));
        Add("PathOrder", typeof(double));

        Add("U", typeof(double)); Add("U_class", typeof(string));
        Add("UCI_D", typeof(double));
        Add("UCI_H", typeof(double));
        Add("UCI", typeof(double));
        Add("UCI_class", typeof(string));

        Add("C", typeof(double));
        Add("F", typeof(double));

        var topH = H.OrderByDescending(v => v).Take(Math.Min(10, n)).ToArray();
        double avgH = topH.Average();

        double LayerCat(double h)
        {
            if (h > avgH * (2.0 / 3.0)) return 1;
            if (h > avgH * (1.0 / 3.0)) return 0;
            return -1;
        }

        for (int i = 0; i < n; i++)
        {
            var r = outT.Rows[i];

            r["Tree1"] = Tree[nn[i, 0]];
            r["Tree2"] = Tree[nn[i, 1]];
            r["Tree3"] = Tree[nn[i, 2]];
            r["Tree4"] = Tree[nn[i, 3]];
            r["Dist1"] = dd[i, 0];
            r["Dist2"] = dd[i, 1];
            r["Dist3"] = dd[i, 2];
            r["Dist4"] = dd[i, 3];

            var ang = new double[4];
            for (int k = 0; k < 4; k++)
            {
                double dx = X[nn[i, k]] - X[i];
                double dy = Y[nn[i, k]] - Y[i];
                double a = Math.Atan2(dy, dx);
                if (a < 0) a += 2 * Math.PI;
                ang[k] = a;
            }
            Array.Sort(ang);
            var gaps = new double[4];
            gaps[0] = ang[1] - ang[0];
            gaps[1] = ang[2] - ang[1];
            gaps[2] = ang[3] - ang[2];
            gaps[3] = 2 * Math.PI - (ang[3] - ang[0]);
            double thr = 72 * Math.PI / 180.0;
            r["W"] = gaps.Count(g => g < thr) / 4.0;

            int diff = 0;
            for (int k = 0; k < 4; k++)
                if (Sp[nn[i, k]] != Sp[i]) diff++;
            r["M"] = diff / 4.0;

            double lc = LayerCat(H[i]);
            double lc1 = LayerCat(H[nn[i, 0]]);
            double lc2 = LayerCat(H[nn[i, 1]]);
            double lc3 = LayerCat(H[nn[i, 2]]);
            double lc4 = LayerCat(H[nn[i, 3]]);
            r["LayerCategory"] = lc;
            r["LayerCategory1"] = lc1;
            r["LayerCategory2"] = lc2;
            r["LayerCategory3"] = lc3;
            r["LayerCategory4"] = lc4;

            var uniq = new HashSet<double> { lc, lc1, lc2, lc3, lc4 };
            r["LayerCount"] = (double)uniq.Count;

            int dis = 0;
            if (lc1 != lc) dis++;
            if (lc2 != lc) dis++;
            if (lc3 != lc) dis++;
            if (lc4 != lc) dis++;
            r["Discreteness"] = dis;
            r["LayerIndex"] = (uniq.Count / 3.0) * (dis / 4.0);

            double openSum = 0;
            for (int k = 0; k < 4; k++)
                openSum += dd[i, k] / Math.Max(H[nn[i, k]], double.Epsilon);
            double openness = openSum / 4.0;
            r["Openness"] = openness;
            r["K"] = openness;

            r["DominantHeight"] = avgH;

            double dbhi = Math.Max(DBH[i], double.Epsilon);
            double ci = 0;
            for (int k = 0; k < 4; k++)
            {
                double dij = Math.Max(dd[i, k], 0.1);
                double cij = DBH[nn[i, k]] / (dbhi * dij);
                r[$"CI{k + 1}"] = cij;
                ci += cij;
            }
            r["CI"] = ci;

            double dbh_cm = DBH[i] * 100.0;
            r["PathOrder"] = 4.0 * Math.Floor((dbh_cm - 2.0) / 4.0) + 4.0;

            int biggerD = 0;
            for (int k = 0; k < 4; k++)
                if (DBH[nn[i, k]] > DBH[i]) biggerD++;
            double U = biggerD / 4.0;
            r["U"] = U;
            r["U_class"] = UClass(U);

            double Di = Math.Max(DBH[i], double.Epsilon);
            double Hi = Math.Max(H[i], double.Epsilon);

            double u_d = 0, u_h = 0;
            for (int k = 0; k < 4; k++)
            {
                int j = nn[i, k];
                double Dj = Math.Max(DBH[j], double.Epsilon);
                double Hj = Math.Max(H[j], double.Epsilon);
                if (Dj > Di) u_d += 1;
                if (Hj > Hi) u_h += 1;
            }
            u_d /= 4.0;
            u_h /= 4.0;

            double sumTermD = 0;
            double sumTermH = 0;

            for (int k = 0; k < 4; k++)
            {
                int j = nn[i, k];
                double dij = Math.Max(dd[i, k], double.Epsilon);

                double Dj = Math.Max(DBH[j], double.Epsilon);
                double Hj = Math.Max(H[j], double.Epsilon);

                double cH = (Hj > Hi) ? 1.0 : 0.0;
                double a1H = Math.Atan(Math.Min(Hi, Hj) / dij);
                double a2H = Math.Max(Math.Atan((Hj - Hi) / dij), 0.0);
                double termH = (a1H + a2H * cH) / Math.PI;
                sumTermH += termH;

                double cD = (Dj > Di) ? 1.0 : 0.0;
                double a1D = Math.Atan(Math.Min(Di, Dj) / dij);
                double a2D = Math.Max(Math.Atan((Dj - Di) / dij), 0.0);
                double termD = (a1D + a2D * cD) / Math.PI;
                sumTermD += termD;
            }

            double UCI_H = Clamp01((sumTermH / 4.0) * u_h);
            double UCI_D = Clamp01((sumTermD / 4.0) * u_d);

            r["UCI_D"] = UCI_D;
            r["UCI_H"] = UCI_H;
            r["UCI"] = UCI_D;
            r["UCI_class"] = UCIClass(UCI_D);

            int inter = 0;
            for (int k = 0; k < 4; k++)
            {
                int j = nn[i, k];
                if (dd[i, k] < (CD[i] + CD[j])) inter++;
            }
            r["C"] = Math.Round((inter / 4.0) * 4.0) / 4.0;

            double rinfl = inflK * CD[i];
            int free = 0;
            for (int k = 0; k < 4; k++)
                if (dd[i, k] > rinfl) free++;
            r["F"] = free / 4.0;
        }

        return outT;
    }

    private static double Clamp01(double x)
    {
        if (double.IsNaN(x) || double.IsInfinity(x)) return 0;
        if (x < 0) return 0;
        if (x >= 1) return 0.999999;
        return x;
    }

    private static string UClass(double u)
    {
        if (u == 0) return "优势";
        if (u <= 0.25) return "亚优势";
        if (u <= 0.50) return "中庸";
        if (u <= 0.75) return "劣势";
        return "绝对劣势";
    }

    private static string UCIClass(double uci)
    {
        if (uci == 0) return "无竞争";
        if (uci <= 0.25) return "较小";
        if (uci <= 0.50) return "中等";
        if (uci <= 0.75) return "较大";
        return "极大";
    }
}

internal sealed class KDTree2D
{
    private readonly (double x, double y)[] _pts;
    private readonly Node? _root;

    private sealed class Node
    {
        public int Index;
        public int Axis;
        public Node? Left;
        public Node? Right;
        public Node(int idx, int axis) { Index = idx; Axis = axis; }
    }

    public KDTree2D((double x, double y)[] pts)
    {
        _pts = pts;
        var idx = Enumerable.Range(0, pts.Length).ToArray();
        _root = Build(idx, 0);
    }

    private Node? Build(int[] idx, int depth)
    {
        if (idx.Length == 0) return null;
        int axis = depth % 2;

        Array.Sort(idx, (a, b) =>
        {
            double va = axis == 0 ? _pts[a].x : _pts[a].y;
            double vb = axis == 0 ? _pts[b].x : _pts[b].y;
            return va.CompareTo(vb);
        });

        int mid = idx.Length / 2;
        var node = new Node(idx[mid], axis);
        node.Left = Build(idx.Take(mid).ToArray(), depth + 1);
        node.Right = Build(idx.Skip(mid + 1).ToArray(), depth + 1);
        return node;
    }

    public (int index, double dist2)[] KNearest(int queryIndex, int k)
    {
        var q = _pts[queryIndex];
        var knn = new FixedKNN(k);
        Search(_root, q, knn);
        return knn.Items();
    }

    private void Search(Node? node, (double x, double y) q, FixedKNN knn)
    {
        if (node is null) return;

        var p = _pts[node.Index];
        double dx = p.x - q.x;
        double dy = p.y - q.y;
        double d2 = dx * dx + dy * dy;
        knn.Consider(node.Index, d2);

        double diff = (node.Axis == 0) ? (q.x - p.x) : (q.y - p.y);
        Node? near = diff <= 0 ? node.Left : node.Right;
        Node? far = diff <= 0 ? node.Right : node.Left;

        Search(near, q, knn);

        double worst = knn.WorstDist2();
        if (!knn.IsFull() || diff * diff < worst)
            Search(far, q, knn);
    }

    private sealed class FixedKNN
    {
        private readonly int _k;
        private readonly List<(int idx, double d2)> _list;

        public FixedKNN(int k) { _k = Math.Max(1, k); _list = new List<(int, double)>(_k); }

        public void Consider(int idx, double d2)
        {
            int pos = _list.FindIndex(t => d2 < t.d2);
            if (pos < 0) _list.Add((idx, d2));
            else _list.Insert(pos, (idx, d2));

            if (_list.Count > _k) _list.RemoveAt(_list.Count - 1);
        }

        public bool IsFull() => _list.Count >= _k;
        public double WorstDist2() => _list.Count == 0 ? double.PositiveInfinity : _list[^1].d2;
        public (int index, double dist2)[] Items() => _list.Select(t => (t.idx, t.d2)).ToArray();
    }
}
'@ | Set-Content -Encoding utf8 .\Services\MetricsEngine.cs

# ---------------- Services: EntropyQ.cs ----------------
@'
using System;
using System.Data;
using System.Globalization;
using System.Linq;

namespace ForestMetricsV5.Services;

public static class EntropyQ
{
    public static (DataTable ew, DataTable qTable, double qAll, double qCore, double qBuf) ComputeEntropyAndQ(DataTable t)
    {
        Require(t, "W","M","LayerIndex","K","F","C","U","UCI","Tree","Zone","IsCore");

        bool[] coreMask = t.AsEnumerable().Select(r => (bool)r["IsCore"]).ToArray();
        if (coreMask.Count(x => x) < 5) coreMask = coreMask.Select(_ => true).ToArray();

        double[][] Xall = new double[t.Rows.Count][];
        for (int i = 0; i < t.Rows.Count; i++)
        {
            var r = t.Rows[i];
            double W = ToD(r["W"]);
            double M = ToD(r["M"]);
            double S = ToD(r["LayerIndex"]);
            double K = ToD(r["K"]);
            double C = ToD(r["C"]);
            double U = ToD(r["U"]);
            double UCI = ToD(r["UCI"]);
            double F = ToD(r["F"]);

            Xall[i] = new[] { Math.Abs(W - 0.5), C, U, UCI, M, S, K, F };
        }

        var Xcore = Xall.Where((row, idx) => coreMask[idx]).ToArray();
        string[] kind = new[] { "neg","neg","neg","neg","pos","pos","pos","pos" };

        var w = EntropyWeight(Xcore, kind);

        double Ew = w[0], Ec = w[1], Eu = w[2], Euci = w[3], Em = w[4], Es = w[5], Ek = w[6], Ef = w[7];

        EnsureCol(t, "Q", typeof(double));

        for (int i = 0; i < t.Rows.Count; i++)
        {
            var r = t.Rows[i];
            double W = ToD(r["W"]);
            double M = ToD(r["M"]);
            double S = ToD(r["LayerIndex"]);
            double K = ToD(r["K"]);
            double C = ToD(r["C"]);
            double U = ToD(r["U"]);
            double UCI = ToD(r["UCI"]);
            double F = ToD(r["F"]);

            double num = ((1 + K) * Ek) * ((1 + M) * Em) * ((1 + S) * Es) * ((1 + F) * Ef);
            double den = ((1 + Math.Abs(W - 0.5)) * Ew) * ((1 + C) * Ec) * ((1 + U) * Eu) * ((1 + UCI) * Euci);

            double Q = (den == 0) ? double.NaN : (num / den);
            r["Q"] = Q;
        }

        double qAll = Mean(t, "Q", _ => true);
        double qCore = Mean(t, "Q", r => (bool)r["IsCore"]);
        double qBuf = Mean(t, "Q", r => !(bool)r["IsCore"]);

        var ew = new DataTable("EntropyWeights");
        ew.Columns.Add("Metric", typeof(string));
        ew.Columns.Add("Weight", typeof(double));
        string[] metric = new[] { "|W-0.5|","C","U","UCI","M","S","K","F" };
        for (int i = 0; i < 8; i++) ew.Rows.Add(metric[i], w[i]);

        var qt = new DataTable("QTable");
        qt.Columns.Add("Tree", typeof(double));
        qt.Columns.Add("Zone", typeof(string));
        qt.Columns.Add("IsCore", typeof(bool));
        qt.Columns.Add("W", typeof(double));
        qt.Columns.Add("M", typeof(double));
        qt.Columns.Add("S", typeof(double));
        qt.Columns.Add("K", typeof(double));
        qt.Columns.Add("C", typeof(double));
        qt.Columns.Add("U", typeof(double));
        qt.Columns.Add("UCI", typeof(double));
        qt.Columns.Add("F", typeof(double));
        qt.Columns.Add("Ew", typeof(double));
        qt.Columns.Add("Em", typeof(double));
        qt.Columns.Add("Es", typeof(double));
        qt.Columns.Add("Ek", typeof(double));
        qt.Columns.Add("Ec", typeof(double));
        qt.Columns.Add("Eu", typeof(double));
        qt.Columns.Add("Euci", typeof(double));
        qt.Columns.Add("Ef", typeof(double));
        qt.Columns.Add("Q", typeof(double));

        for (int i = 0; i < t.Rows.Count; i++)
        {
            var r = t.Rows[i];
            qt.Rows.Add(
                ToD(r["Tree"]), r["Zone"]?.ToString() ?? "", (bool)r["IsCore"],
                ToD(r["W"]), ToD(r["M"]), ToD(r["LayerIndex"]), ToD(r["K"]),
                ToD(r["C"]), ToD(r["U"]), ToD(r["UCI"]), ToD(r["F"]),
                Ew, Em, Es, Ek, Ec, Eu, Euci, Ef,
                ToD(r["Q"])
            );
        }

        return (ew, qt, qAll, qCore, qBuf);
    }

    private static double[] EntropyWeight(double[][] X, string[] kind)
    {
        int n = X.Length;
        int m = X[0].Length;

        double[][] Y = new double[n][];
        for (int i = 0; i < n; i++) Y[i] = (double[])X[i].Clone();

        for (int j = 0; j < m; j++)
        {
            double min = Y.Min(r => r[j]);
            double max = Y.Max(r => r[j]);
            double rng = max - min;

            for (int i = 0; i < n; i++)
            {
                double v = (rng == 0 || !double.IsFinite(rng)) ? 0 : (Y[i][j] - min) / rng;
                if (kind[j] == "neg") v = 1 - v;
                Y[i][j] = v;
            }
        }

        double[] colSum = new double[m];
        for (int j = 0; j < m; j++)
        {
            colSum[j] = 0;
            for (int i = 0; i < n; i++) colSum[j] += Y[i][j];
            if (colSum[j] == 0) colSum[j] = double.Epsilon;
        }

        double[,] P = new double[n, m];
        for (int i = 0; i < n; i++)
            for (int j = 0; j < m; j++)
                P[i, j] = Y[i][j] / colSum[j];

        double k = 1.0 / Math.Log(n);
        double[] E = new double[m];
        for (int j = 0; j < m; j++)
        {
            double s = 0;
            for (int i = 0; i < n; i++)
                s += P[i, j] * Math.Log(P[i, j] + double.Epsilon);
            E[j] = -k * s;
        }

        double[] d = E.Select(e => 1 - e).ToArray();
        double sumd = d.Sum();
        if (sumd == 0 || !double.IsFinite(sumd))
        {
            for (int j = 0; j < m; j++) d[j] = 1;
            sumd = d.Sum();
        }

        return d.Select(x => x / sumd).ToArray();
    }

    private static void Require(DataTable t, params string[] cols)
    {
        foreach (var c in cols)
            if (!t.Columns.Contains(c))
                throw new Exception("EntropyQ missing column: " + c);
    }

    private static double ToD(object? o)
    {
        if (o is null || o is DBNull) return double.NaN;
        if (o is double d) return d;
        if (double.TryParse(o.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out d)) return d;
        if (double.TryParse(o.ToString(), NumberStyles.Any, CultureInfo.CurrentCulture, out d)) return d;
        return double.NaN;
    }

    private static void EnsureCol(DataTable t, string name, Type type)
    {
        if (!t.Columns.Contains(name)) t.Columns.Add(name, type);
    }

    private static double Mean(DataTable t, string col, Func<DataRow, bool> mask)
    {
        var arr = t.AsEnumerable().Where(mask).Select(r => ToD(r[col])).Where(x => double.IsFinite(x)).ToArray();
        return arr.Length == 0 ? double.NaN : arr.Average();
    }
}
'@ | Set-Content -Encoding utf8 .\Services\EntropyQ.cs

# ---------------- Services: StandMeans.cs ----------------
@'
using System;
using System.Data;
using System.Globalization;
using System.Linq;

namespace ForestMetricsV5.Services;

public static class StandMeans
{
    public static DataTable BuildStandMeans(DataTable t, double qAll, double qCore, double qBuf)
    {
        var summary = new DataTable("StandMeans");
        summary.Columns.Add("Group", typeof(string));
        summary.Columns.Add("Metric", typeof(string));
        summary.Columns.Add("Mean", typeof(double));

        string[] cands = new[]
        {
            "W","M","U","UCI","C","F","Openness","K","CI","LayerIndex","UCI_D","UCI_H","DistToBoundary","Q"
        };

        var present = cands.Where(c => t.Columns.Contains(c)).ToArray();

        AddGroup("ALL",    r => true);
        AddGroup("CORE",   r => (t.Columns.Contains("IsCore") && (bool)r["IsCore"]));
        AddGroup("BUFFER", r => (t.Columns.Contains("IsCore") && !(bool)r["IsCore"]));

        summary.Rows.Add("ALL", "Qbar", qAll);
        summary.Rows.Add("CORE", "Qbar", qCore);
        summary.Rows.Add("BUFFER", "Qbar", qBuf);

        return summary;

        void AddGroup(string g, Func<DataRow, bool> mask)
        {
            foreach (var m in present)
            {
                var arr = t.AsEnumerable().Where(mask).Select(r => ToD(r[m])).Where(x => double.IsFinite(x)).ToArray();
                if (arr.Length == 0) continue;
                summary.Rows.Add(g, m, arr.Average());
            }
        }

        static double ToD(object? o)
        {
            if (o is null || o is DBNull) return double.NaN;
            if (o is double d) return d;
            if (double.TryParse(o.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out d)) return d;
            if (double.TryParse(o.ToString(), NumberStyles.Any, CultureInfo.CurrentCulture, out d)) return d;
            return double.NaN;
        }
    }
}
'@ | Set-Content -Encoding utf8 .\Services\StandMeans.cs

# ---------------- Services: Exporter.cs ----------------
@'
using System;
using System.Data;
using System.IO;
using System.Text;

namespace ForestMetricsV5.Services;

public static class Exporter
{
    public record ExportPaths(string? Main, string? Means, string? Entropy, string? QTree);

    public static ExportPaths WriteAll(
        DataTable main, DataTable means, DataTable entropy, DataTable qtree,
        string inputPath, string outDir, string format,
        bool writeMain, bool writeMeans, bool writeEntropy, bool writeQ)
    {
        var inFolder = Path.GetDirectoryName(inputPath) ?? ".";
        var baseName = Path.GetFileNameWithoutExtension(inputPath);
        var ext = Path.GetExtension(inputPath);

        if (string.IsNullOrWhiteSpace(outDir)) outDir = inFolder;
        Directory.CreateDirectory(outDir);

        bool csv = format.Equals("CSV", StringComparison.OrdinalIgnoreCase);
        var suf = csv ? ".csv" : ".txt";

        string? pMain = null, pMeans = null, pEntropy = null, pQ = null;

        if (writeMain)
        {
            pMain = Path.Combine(outDir, $"{baseName}{ext}-输出{suf}");
            WriteTable(main, pMain, csv);
        }
        if (writeMeans)
        {
            pMeans = Path.Combine(outDir, $"{baseName}{ext}-输出-林分均值{suf}");
            WriteTable(means, pMeans, csv);
        }
        if (writeEntropy)
        {
            pEntropy = Path.Combine(outDir, $"{baseName}{ext}-输出-熵权{suf}");
            WriteTable(entropy, pEntropy, csv);
        }
        if (writeQ)
        {
            pQ = Path.Combine(outDir, $"{baseName}{ext}-输出-综合指数{suf}");
            WriteTable(qtree, pQ, csv);
        }

        return new ExportPaths(pMain, pMeans, pEntropy, pQ);
    }

    private static void WriteTable(DataTable table, string path, bool csv)
    {
        var sep = csv ? "," : "\t";
        using var sw = new StreamWriter(path, false, new UTF8Encoding(true));

        for (int c = 0; c < table.Columns.Count; c++)
        {
            if (c > 0) sw.Write(sep);
            sw.Write(table.Columns[c].ColumnName);
        }
        sw.WriteLine();

        foreach (DataRow r in table.Rows)
        {
            for (int c = 0; c < table.Columns.Count; c++)
            {
                if (c > 0) sw.Write(sep);
                var v = r[c];
                if (v is null || v is DBNull) { sw.Write(""); continue; }

                var s = v.ToString() ?? "";
                if (csv && (s.Contains('"') || s.Contains(',') || s.Contains('\n') || s.Contains('\r')))
                    s = "\"" + s.Replace("\"", "\"\"") + "\"";

                sw.Write(s);
            }
            sw.WriteLine();
        }
    }
}
'@ | Set-Content -Encoding utf8 .\Services\Exporter.cs

# ---------------- Services: ColumnFilter.cs ----------------
@'
using System.Collections.Generic;
using System.Data;

namespace ForestMetricsV5.Services;

public static class ColumnFilter
{
    public static DataTable BuildMainForPreview(DataTable t, bool useWM, bool useLayer, bool useUUCI, bool useCF, bool useQ)
    {
        var keep = new List<string>();

        void Add(params string[] cols)
        {
            foreach (var c in cols)
                if (t.Columns.Contains(c) && !keep.Contains(c))
                    keep.Add(c);
        }

        Add("Tree","X","Y","DBH","Species","Height","CrownRadius");
        Add("BufferSize_m","DistToBoundary","IsCore","Zone");
        Add("Tree1","Tree2","Tree3","Tree4","Dist1","Dist2","Dist3","Dist4");

        if (useWM) Add("W","M");
        if (useLayer) Add("DominantHeight","LayerCategory","LayerCount","LayerIndex","Discreteness","Openness","K","CI1","CI2","CI3","CI4","CI","PathOrder");
        if (useUUCI) Add("U","U_class","UCI","UCI_class","UCI_D","UCI_H");
        if (useCF) Add("C","F");
        if (useQ) Add("Q");

        var outT = new DataTable("Main");
        foreach (var c in keep) outT.Columns.Add(c, t.Columns[c].DataType);

        foreach (DataRow r in t.Rows)
        {
            var nr = outT.NewRow();
            foreach (var c in keep) nr[c] = r[c];
            outT.Rows.Add(nr);
        }
        return outT;
    }
}
'@ | Set-Content -Encoding utf8 .\Services\ColumnFilter.cs

# ---------------- Services: MetricCatalog.cs ----------------
@'
using System.Collections.Generic;

namespace ForestMetricsV5.Services;

public static class MetricCatalog
{
    public static List<string> BuildMetricList(bool useWM, bool useLayer, bool useUUCI, bool useCF, bool useQ)
    {
        var list = new List<string>();
        if (useWM) { list.Add("W"); list.Add("M"); }
        if (useUUCI) { list.Add("U"); list.Add("UCI"); list.Add("UCI_D"); list.Add("UCI_H"); }
        if (useCF) { list.Add("C"); list.Add("F"); }
        if (useLayer) { list.Add("Openness"); list.Add("K"); list.Add("LayerIndex"); list.Add("CI"); list.Add("PathOrder"); }
        if (useQ) { list.Add("Q"); }
        list.Add("DBH");
        list.Add("Height");
        list.Add("DistToBoundary");
        return list;
    }
}
'@ | Set-Content -Encoding utf8 .\Services\MetricCatalog.cs

# ---------------- Services: HistPlotter.cs ----------------
@'
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Series;

namespace ForestMetricsV5.Services;

public static class HistPlotter
{
    public static PlotModel BuildMetricHistogram(DataTable t, string metric, string group)
    {
        var (vals, nTotal) = GetValues(t, metric, group);

        var pm = new PlotModel
        {
            Title = $"{metric} frequency ({group}) | n={vals.Count}",
            TitleFontSize = 20,
            DefaultFontSize = 13,
            Background = OxyColors.White,
            PlotAreaBackground = OxyColors.White
        };

        if (vals.Count == 0)
        {
            pm.Subtitle = "No valid data.";
            return pm;
        }

        bool looksQuarter =
            vals.All(x => x >= -1e-9 && x <= 1.000001) &&
            vals.All(x => Math.Abs(x - Math.Round(x * 4.0) / 4.0) < 1e-9);

        if (looksQuarter)
            return BuildQuarterBars(pm, metric, vals);

        return BuildAutoBars(pm, metric, vals);
    }

    private static PlotModel BuildQuarterBars(PlotModel pm, string metric, List<double> vals)
    {
        double[] cats = { 0.0, 0.25, 0.5, 0.75, 1.0 };
        var counts = new int[cats.Length];

        foreach (var x in vals)
        {
            var q = Math.Round(x * 4.0) / 4.0;
            for (int i = 0; i < cats.Length; i++)
                if (Math.Abs(q - cats[i]) < 1e-9) { counts[i]++; break; }
        }

        pm.Axes.Add(new LinearAxis
        {
            Position = AxisPosition.Bottom,
            Title = metric,
            Minimum = -0.05,
            Maximum = 1.05,
            MajorStep = 0.25,
            MinorStep = 0.05,
            MajorGridlineStyle = LineStyle.Solid,
            MinorGridlineStyle = LineStyle.Dot,
            MajorGridlineColor = OxyColor.FromAColor(60, OxyColors.Gray),
            MinorGridlineColor = OxyColor.FromAColor(30, OxyColors.Gray)
        });

        pm.Axes.Add(new LinearAxis
        {
            Position = AxisPosition.Left,
            Title = "Count",
            Minimum = 0,
            MajorGridlineStyle = LineStyle.Solid,
            MinorGridlineStyle = LineStyle.Dot,
            MajorGridlineColor = OxyColor.FromAColor(60, OxyColors.Gray),
            MinorGridlineColor = OxyColor.FromAColor(30, OxyColors.Gray)
        });

        var bars = new RectangleBarSeries
        {
            StrokeThickness = 1.2,
            FillColor = OxyColor.FromAColor(120, OxyColors.SeaGreen),
            StrokeColor = OxyColors.ForestGreen
        };

        double w = 0.18;
        for (int i = 0; i < cats.Length; i++)
        {
            double x0 = cats[i] - w / 2;
            double x1 = cats[i] + w / 2;
            bars.Items.Add(new RectangleBarItem(x0, 0, x1, counts[i]));
        }

        pm.Series.Add(bars);

        var med = vals.OrderBy(x => x).ElementAt(vals.Count / 2);
        pm.Subtitle = $"mean={vals.Average():0.###}, median={med:0.###}";
        return pm;
    }

    private static PlotModel BuildAutoBars(PlotModel pm, string metric, List<double> vals)
    {
        double vmin = vals.Min();
        double vmax = vals.Max();
        if (vmax <= vmin) vmax = vmin + 1e-6;

        int bins = 14;
        double step = (vmax - vmin) / bins;
        var counts = new int[bins];

        foreach (var x in vals)
        {
            double xx = Math.Min(vmax - 1e-12, Math.Max(vmin, x));
            int b = (int)Math.Floor((xx - vmin) / step);
            b = Math.Max(0, Math.Min(bins - 1, b));
            counts[b]++;
        }

        pm.Axes.Add(new LinearAxis
        {
            Position = AxisPosition.Bottom,
            Title = "Value",
            Minimum = vmin,
            Maximum = vmax,
            MajorGridlineStyle = LineStyle.Solid,
            MinorGridlineStyle = LineStyle.Dot,
            MajorGridlineColor = OxyColor.FromAColor(60, OxyColors.Gray),
            MinorGridlineColor = OxyColor.FromAColor(30, OxyColors.Gray)
        });

        pm.Axes.Add(new LinearAxis
        {
            Position = AxisPosition.Left,
            Title = "Count",
            Minimum = 0,
            MajorGridlineStyle = LineStyle.Solid,
            MinorGridlineStyle = LineStyle.Dot,
            MajorGridlineColor = OxyColor.FromAColor(60, OxyColors.Gray),
            MinorGridlineColor = OxyColor.FromAColor(30, OxyColors.Gray)
        });

        var bars = new RectangleBarSeries
        {
            StrokeThickness = 1.0,
            FillColor = OxyColor.FromAColor(110, OxyColors.SteelBlue),
            StrokeColor = OxyColors.DarkSlateGray
        };

        for (int i = 0; i < bins; i++)
        {
            double a = vmin + i * step;
            double b = a + step;
            bars.Items.Add(new RectangleBarItem(a, 0, b, counts[i]));
        }

        pm.Series.Add(bars);

        var med = vals.OrderBy(x => x).ElementAt(vals.Count / 2);
        pm.Subtitle = $"mean={vals.Average():0.###}, median={med:0.###}";
        return pm;
    }

    private static (List<double> vals, int n) GetValues(DataTable t, string metric, string group)
    {
        if (!t.Columns.Contains(metric)) return (new List<double>(), 0);

        bool useCore = group.Equals("CORE", StringComparison.OrdinalIgnoreCase);
        bool useBuf  = group.Equals("BUFFER", StringComparison.OrdinalIgnoreCase);

        var list = new List<double>(t.Rows.Count);
        int n = 0;

        foreach (DataRow r in t.Rows)
        {
            if (useCore && t.Columns.Contains("IsCore") && !(bool)r["IsCore"]) continue;
            if (useBuf  && t.Columns.Contains("IsCore") &&  (bool)r["IsCore"]) continue;

            var o = r[metric];
            if (o is null || o is DBNull) continue;

            if (double.TryParse(o.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out var x) ||
                double.TryParse(o.ToString(), NumberStyles.Any, CultureInfo.CurrentCulture, out x))
            {
                if (double.IsFinite(x)) list.Add(x);
            }
            n++;
        }

        return (list, n);
    }
}
'@ | Set-Content -Encoding utf8 .\Services\HistPlotter.cs

# ---------------- Services: VoronoiPlotter.cs ----------------
@'
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Annotations;
using OxyPlot.Series;

namespace ForestMetricsV5.Services;

public static class VoronoiPlotter
{
    public static PlotModel BuildVoronoiApproxModel(DataTable t, string title)
    {
        var pm = new PlotModel
        {
            Title = title,
            TitleFontSize = 20,
            DefaultFontSize = 12,
            Background = OxyColors.White,
            PlotAreaBackground = OxyColors.White
        };

        if (!t.Columns.Contains("X") || !t.Columns.Contains("Y")) return pm;

        var xs = t.AsEnumerable().Select(r => (double)r["X"]).ToArray();
        var ys = t.AsEnumerable().Select(r => (double)r["Y"]).ToArray();

        double xmin = xs.Min(), xmax = xs.Max();
        double ymin = ys.Min(), ymax = ys.Max();

        pm.Axes.Add(new LinearAxis
        {
            Position = AxisPosition.Bottom,
            Title = "X",
            Minimum = xmin,
            Maximum = xmax,
            MajorGridlineStyle = LineStyle.Solid,
            MinorGridlineStyle = LineStyle.Dot,
            MajorGridlineColor = OxyColor.FromAColor(60, OxyColors.Gray),
            MinorGridlineColor = OxyColor.FromAColor(30, OxyColors.Gray)
        });

        pm.Axes.Add(new LinearAxis
        {
            Position = AxisPosition.Left,
            Title = "Y",
            Minimum = ymin,
            Maximum = ymax,
            MajorGridlineStyle = LineStyle.Solid,
            MinorGridlineStyle = LineStyle.Dot,
            MajorGridlineColor = OxyColor.FromAColor(60, OxyColors.Gray),
            MinorGridlineColor = OxyColor.FromAColor(30, OxyColors.Gray)
        });

        pm.Annotations.Add(new RectangleAnnotation
        {
            MinimumX = xmin, MaximumX = xmax, MinimumY = ymin, MaximumY = ymax,
            Stroke = OxyColors.ForestGreen,
            StrokeThickness = 2.0,
            Fill = OxyColors.Undefined
        });

        if (t.Columns.Contains("BufferSize_m"))
        {
            double b = t.AsEnumerable().Select(r => (double)r["BufferSize_m"]).FirstOrDefault();
            if (b > 0)
            {
                pm.Annotations.Add(new RectangleAnnotation
                {
                    MinimumX = xmin + b, MaximumX = xmax - b,
                    MinimumY = ymin + b, MaximumY = ymax - b,
                    Stroke = OxyColors.OrangeRed,
                    StrokeThickness = 2.0,
                    Fill = OxyColors.Undefined
                });
            }
        }

        var seg = BuildRasterBoundarySegments(t, xmin, xmax, ymin, ymax);
        if (seg.Count > 0)
        {
            var ls = new LineSeries
            {
                StrokeThickness = 1.0,
                Color = OxyColor.FromAColor(140, OxyColors.Gray)
            };

            foreach (var s in seg)
            {
                ls.Points.Add(new DataPoint(s.x1, s.y1));
                ls.Points.Add(new DataPoint(s.x2, s.y2));
                ls.Points.Add(DataPoint.Undefined);
            }
            pm.Series.Add(ls);
        }

        bool hasCore = t.Columns.Contains("IsCore");

        var core = new ScatterSeries
        {
            MarkerType = MarkerType.Circle,
            MarkerSize = 3.4,
            MarkerFill = OxyColors.Red,
            MarkerStroke = OxyColors.DarkRed,
            MarkerStrokeThickness = 0.6
        };

        var buf = new ScatterSeries
        {
            MarkerType = MarkerType.Circle,
            MarkerSize = 3.4,
            MarkerFill = OxyColors.RoyalBlue,
            MarkerStroke = OxyColors.DarkBlue,
            MarkerStrokeThickness = 0.6
        };

        for (int i = 0; i < t.Rows.Count; i++)
        {
            var r = t.Rows[i];
            double x = (double)r["X"], y = (double)r["Y"];
            bool isCore = hasCore ? (bool)r["IsCore"] : true;

            if (isCore) core.Points.Add(new ScatterPoint(x, y));
            else buf.Points.Add(new ScatterPoint(x, y));
        }

        pm.Series.Add(core);
        pm.Series.Add(buf);

        return pm;
    }

    private static List<(double x1, double y1, double x2, double y2)> BuildRasterBoundarySegments(
        DataTable t, double xmin, double xmax, double ymin, double ymax)
    {
        int n = t.Rows.Count;
        int res = n <= 200 ? 240 : (n <= 800 ? 170 : 120);
        if (n > 4000) res = 100;

        var pts = t.AsEnumerable()
            .Select(r => ((double)r["X"], (double)r["Y"]))
            .ToArray();
        if (pts.Length < 2) return new();

        double dx = (xmax - xmin) / (res - 1);
        double dy = (ymax - ymin) / (res - 1);
        if (dx <= 0 || dy <= 0) return new();

        int[,] lab = new int[res, res];

        for (int j = 0; j < res; j++)
        {
            double y = ymin + j * dy;
            for (int i = 0; i < res; i++)
            {
                double x = xmin + i * dx;

                double best = double.PositiveInfinity;
                int bestId = 0;
                for (int p = 0; p < pts.Length; p++)
                {
                    double ddx = pts[p].Item1 - x;
                    double ddy = pts[p].Item2 - y;
                    double d2 = ddx * ddx + ddy * ddy;
                    if (d2 < best) { best = d2; bestId = p; }
                }
                lab[i, j] = bestId;
            }
        }

        var segs = new List<(double x1, double y1, double x2, double y2)>(5000);

        for (int i = 0; i < res - 1; i++)
        {
            double xmid = xmin + (i + 0.5) * dx;
            int j = 0;
            while (j < res)
            {
                bool isB = lab[i, j] != lab[i + 1, j];
                if (!isB) { j++; continue; }

                int j0 = j;
                while (j < res && lab[i, j] != lab[i + 1, j]) j++;
                int j1 = j - 1;

                double y1 = ymin + j0 * dy - dy / 2;
                double y2 = ymin + j1 * dy + dy / 2;
                y1 = Math.Max(ymin, y1);
                y2 = Math.Min(ymax, y2);

                segs.Add((xmid, y1, xmid, y2));
            }
        }

        for (int j = 0; j < res - 1; j++)
        {
            double ymid = ymin + (j + 0.5) * dy;
            int i = 0;
            while (i < res)
            {
                bool isB = lab[i, j] != lab[i, j + 1];
                if (!isB) { i++; continue; }

                int i0 = i;
                while (i < res && lab[i, j] != lab[i, j + 1]) i++;
                int i1 = i - 1;

                double x1 = xmin + i0 * dx - dx / 2;
                double x2 = xmin + i1 * dx + dx / 2;
                x1 = Math.Max(xmin, x1);
                x2 = Math.Min(xmax, x2);

                segs.Add((x1, ymid, x2, ymid));
            }
        }

        return segs;
    }
}
'@ | Set-Content -Encoding utf8 .\Services\VoronoiPlotter.cs

# ---------------- Services: PlotExport.cs ----------------
@'
using System;
using System.IO;
using OxyPlot;

namespace ForestMetricsV5.Services;

public static class PlotExport
{
    public static bool TryExportPng(IPlotModel model, string path, int width, int height)
    {
        try
        {
            var t = Type.GetType("OxyPlot.SkiaSharp.PngExporter, OxyPlot.SkiaSharp");
            if (t == null) return false;

            var exporter = Activator.CreateInstance(t);
            t.GetProperty("Width")?.SetValue(exporter, width);
            t.GetProperty("Height")?.SetValue(exporter, height);

            using var fs = File.Create(path);

            var m = t.GetMethod("Export", new[] { typeof(IPlotModel), typeof(Stream) });
            if (m == null) return false;

            m.Invoke(exporter, new object[] { model, fs });
            return true;
        }
        catch
        {
            return false;
        }
    }
}
'@ | Set-Content -Encoding utf8 .\Services\PlotExport.cs

# ======================================================================
# [6/10] CLEAN + RESTORE + BUILD
# ======================================================================
Write-Host "[6/10] Clean bin/obj ..." -ForegroundColor Cyan
Remove-Item -Recurse -Force .\bin,.\obj -ErrorAction SilentlyContinue

Write-Host "[7/10] dotnet restore ..." -ForegroundColor Cyan
dotnet restore

Write-Host "[8/10] dotnet build ..." -ForegroundColor Cyan
dotnet build

# ======================================================================
# [9/10] PUBLISH SELF-CONTAINED SINGLE-FILE EXE
# ======================================================================
Write-Host "[9/10] dotnet publish (self-contained single-file) ..." -ForegroundColor Cyan
dotnet publish -c Release -r win-x64 --self-contained true `
  /p:PublishSingleFile=true `
  /p:IncludeNativeLibrariesForSelfExtract=true `
  /p:DebugType=None /p:DebugSymbols=false

# ======================================================================
# [10/10] RENAME/COPY EXE TO RELEASE FOLDER (TEAM BRAND)
# ======================================================================
$pub = Join-Path $projDir "bin\Release\net8.0\win-x64\publish"
$srcExe = Join-Path $pub "${APP}.exe"
if (-not (Test-Path $srcExe)) {
  $maybe = Get-ChildItem $pub -Filter "*.exe" | Select-Object -First 1
  if ($maybe) { $srcExe = $maybe.FullName }
}
if (-not (Test-Path $srcExe)) { throw "publish 目录未找到 exe：$pub" }

$dstExe = Join-Path $releaseDir $EXE_NAME
Copy-Item -Force $srcExe $dstExe

Write-Host ""
Write-Host "== RELEASE OUTPUT ==" -ForegroundColor Green
Write-Host "EXE: $dstExe" -ForegroundColor Green
Write-Host ""
Write-Host "调试运行：cd $projDir ; dotnet run" -ForegroundColor Yellow
Write-Host "发布完成：$dstExe（可拷贝到其它 Win10/11 x64 电脑直接运行）" -ForegroundColor Yellow