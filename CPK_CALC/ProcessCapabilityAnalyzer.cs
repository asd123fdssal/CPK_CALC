// WinForms / Charting
// MigraDocCore (문자열 경로 AddImage 버전)
using MigraDocCore.DocumentObjectModel;
using MigraDocCore.DocumentObjectModel.MigraDoc.DocumentObjectModel.Shapes;
using MigraDocCore.Rendering;
using System.Text;
using System.Windows.Forms.DataVisualization.Charting;
// System.Drawing은 SD로 별칭
using SD = System.Drawing;
using WFBorderStyle = System.Windows.Forms.BorderStyle;
using WFOrient = System.Windows.Forms.Orientation;

namespace CpkTool
{
    // ----------------------- 데이터 파싱 -----------------------
    public class TestDataParser
    {
        public class TestItem
        {
            public string Step { get; set; } = "";
            public string ItemName { get; set; } = "";
            public string Spec { get; set; } = "";
            public double[] Values { get; set; } = Array.Empty<double>();
            public string Unit { get; set; } = "";
            public double LSL { get; set; }
            public double USL { get; set; }
            public string OriginalSpec { get; set; } = "";
        }

        public static List<TestItem> ParseFile(string filePath)
        {
            var testItems = new List<TestItem>();
            var lines = File.ReadAllLines(filePath);
            string currentStep = "";

            foreach (var raw in lines)
            {
                var line = raw?.Trim();
                if (string.IsNullOrWhiteSpace(line) ||
                    line.StartsWith("STEP ITEM") ||
                    line.Contains("[")) continue;

                var testItem = ParseLine(line, ref currentStep);
                if (testItem != null && testItem.Values.Any(v => v != 0))
                    testItems.Add(testItem);
            }

            return testItems;
        }

        private static TestItem? ParseLine(string line, ref string currentStep)
        {
            try
            {
                var parts = line.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length < 4) return null;

                var testItem = new TestItem();

                if (char.IsDigit(parts[0][0]) || parts[0].Contains("."))
                {
                    currentStep = parts[0];
                    testItem.Step = currentStep;

                    int specIndex = -1;
                    for (int i = 1; i < parts.Length; i++)
                        if (IsSpecFormat(parts[i])) { specIndex = i; break; }
                    if (specIndex == -1) return null;

                    testItem.ItemName = string.Join(" ", parts, 1, specIndex - 1);
                    testItem.Spec = parts[specIndex];
                    testItem.OriginalSpec = testItem.Spec;

                    ParseSpecification(testItem.Spec, out double lsl, out double usl);
                    testItem.LSL = lsl; testItem.USL = usl;

                    var values = new List<double>();
                    int unitIndex = parts.Length - 1;
                    for (int i = specIndex + 1; i < unitIndex; i++)
                        if (double.TryParse(parts[i], out double v)) values.Add(v);

                    testItem.Values = values.ToArray();
                    testItem.Unit = parts[unitIndex];
                }
                else
                {
                    testItem.Step = currentStep;

                    int specIndex = -1;
                    for (int i = 0; i < parts.Length; i++)
                        if (IsSpecFormat(parts[i])) { specIndex = i; break; }
                    if (specIndex == -1) return null;

                    testItem.ItemName = string.Join(" ", parts, 0, specIndex);
                    testItem.Spec = parts[specIndex];
                    testItem.OriginalSpec = testItem.Spec;

                    ParseSpecification(testItem.Spec, out double lsl, out double usl);
                    testItem.LSL = lsl; testItem.USL = usl;

                    var values = new List<double>();
                    int unitIndex = parts.Length - 1;
                    for (int i = specIndex + 1; i < unitIndex; i++)
                        if (double.TryParse(parts[i], out double v)) values.Add(v);

                    testItem.Values = values.ToArray();
                    testItem.Unit = parts[unitIndex];
                }

                return testItem;
            }
            catch { return null; }
        }

        private static bool IsSpecFormat(string text)
            => text.Contains("~")
               && text.Split('~').Length == 2
               && text.Split('~').All(part => double.TryParse(part, out _));

        private static void ParseSpecification(string spec, out double lsl, out double usl)
        {
            lsl = double.MinValue; usl = double.MaxValue;
            if (!spec.Contains("~")) return;
            var parts = spec.Split('~');
            if (parts.Length != 2) return;
            double.TryParse(parts[0], out lsl);
            double.TryParse(parts[1], out usl);
        }
    }

    // ----------------------- 공정능력 계산 -----------------------
    public class ProcessCapabilityAnalyzer
    {
        public class CapabilityResult
        {
            public double Cp { get; set; }
            public double Cpk { get; set; }
            public double CpkLower { get; set; }
            public double CpkUpper { get; set; }
            public double Pp { get; set; }
            public double Ppk { get; set; }
            public double Mean { get; set; }
            public double StdDev { get; set; }
            public double LSL { get; set; }
            public double USL { get; set; }
            public int SampleSize { get; set; }
            public bool IsCapable { get; set; }
            public double[] RawData { get; set; } = Array.Empty<double>();

            public string Step { get; set; } = "";  
            public string ItemName { get; set; } = "";
            public string Unit { get; set; } = "";
        }

        public static CapabilityResult CalculateProcessCapability(
            IEnumerable<double> data, double lsl, double usl,
            string itemName = "", string unit = "", bool useUnbiasedStdDev = true, string step = "")
        {
            if (data == null || !data.Any()) throw new ArgumentException("데이터가 비어있습니다.");
            if (usl <= lsl) throw new ArgumentException("상한규격이 하한규격보다 작거나 같습니다.");

            var arr = data.ToArray(); int n = arr.Length;
            if (n < 2) throw new ArgumentException("최소 2개 이상의 데이터가 필요합니다.");

            double mean = arr.Average();
            double var = arr.Sum(x => Math.Pow(x - mean, 2)) / (useUnbiasedStdDev ? n - 1 : n);
            double std = Math.Sqrt(var);

            double cp = (usl - lsl) / (6 * std);
            double cpl = (mean - lsl) / (3 * std);
            double cpu = (usl - mean) / (3 * std);
            double cpk = Math.Min(cpl, cpu);

            double totalStd = Math.Sqrt(arr.Sum(x => Math.Pow(x - mean, 2)) / n);
            double pp = (usl - lsl) / (6 * totalStd);
            double ppk = Math.Min((mean - lsl) / (3 * totalStd), (usl - mean) / (3 * totalStd));

            return new CapabilityResult
            {
                Cp = cp,
                Cpk = cpk,
                CpkLower = cpl,
                CpkUpper = cpu,
                Pp = pp,
                Ppk = ppk,
                Mean = mean,
                StdDev = std,
                LSL = lsl,
                USL = usl,
                SampleSize = n,
                IsCapable = cpk >= 1.33,
                RawData = arr,
                ItemName = itemName,
                Unit = unit,
                Step = step
            };
        }
    }

    // ----------------------- 차트 생성 -----------------------
    public class CPKChartGenerator
    {
        public static Chart CreateCapabilityHistogram(
    ProcessCapabilityAnalyzer.CapabilityResult r, int binCount = 20)
        {
            var chart = new Chart { Size = new SD.Size(800, 450) };

            var area = new ChartArea("MainArea")
            {
                AxisX = { Title = $"측정값 ({r.Unit})", MajorGrid = { Enabled = true } },
                AxisY = { Title = "빈도", MajorGrid = { Enabled = true } }
            };
            chart.ChartAreas.Add(area);

            if (r.StdDev < 1e-10 || r.RawData.Length == 0)
            {
                var t = new Title("모든 데이터가 동일하여 히스토그램을 생성할 수 없습니다")
                { Font = new SD.Font("Arial", 12, SD.FontStyle.Bold) };
                chart.Titles.Add(t);
                return chart;
            }

            var histogram = CreateHistogram(r.RawData, binCount);

            var hist = new Series("히스토그램")
            {
                ChartType = SeriesChartType.Column,
                Color = SD.Color.LightBlue,
                BorderColor = SD.Color.DarkBlue,
                BorderWidth = 1
            };
            foreach (var bin in histogram)
                hist.Points.AddXY(bin.Key, bin.Value);
            chart.Series.Add(hist);

            // -------------------------------
            // ★ LSL / USL 포함하여 전체 범위 계산
            // -------------------------------
            double minVal = Math.Min(r.RawData.Min(), r.LSL);
            double maxVal = Math.Max(r.RawData.Max(), r.USL);
            double padding = (maxVal - minVal) * 0.05;   // 좌우 5% 여유

            area.AxisX.Minimum = minVal - padding;
            area.AxisX.Maximum = maxVal + padding;

            // -------------------------------
            // 정규분포 곡선 추가
            // -------------------------------
            if (r.StdDev > 1e-6)
            {
                var normal = new Series("정규분포")
                {
                    ChartType = SeriesChartType.Spline,
                    Color = SD.Color.Red,
                    BorderWidth = 2
                };

                double range = maxVal - minVal;
                double step = range / 200;
                double binWidth = range / binCount;
                double scale = r.RawData.Length * binWidth;

                for (double x = area.AxisX.Minimum; x <= area.AxisX.Maximum; x += step)
                {
                    double y = NormalPDF(x, r.Mean, r.StdDev) * scale;
                    if (!double.IsNaN(y) && !double.IsInfinity(y))
                        normal.Points.AddXY(x, y);
                }
                chart.Series.Add(normal);
            }

            // 규격선 추가
            AddSpecificationLines(chart, r);

            chart.Titles.Add(new Title($"{r.ItemName} - Cp={r.Cp:F3}, Cpk={r.Cpk:F3}")
            { Font = new SD.Font("Arial", 10, SD.FontStyle.Bold) });
            chart.Legends.Add(new Legend("Legend") { Docking = Docking.Right });
            return chart;
        }


        public static Chart CreateControlChart(ProcessCapabilityAnalyzer.CapabilityResult r)
        {
            var chart = new Chart { Size = new SD.Size(800, 400) };

            var area = new ChartArea("MainArea")
            {
                AxisX = { Title = "샘플 번호", MajorGrid = { Enabled = true } },
                AxisY = { Title = $"측정값 ({r.Unit})", MajorGrid = { Enabled = true } }
            };
            chart.ChartAreas.Add(area);

            var data = new Series("측정값")
            {
                ChartType = SeriesChartType.Line,
                MarkerStyle = MarkerStyle.Circle,
                MarkerSize = 5,
                Color = SD.Color.Blue,
                BorderWidth = 2
            };
            for (int i = 0; i < r.RawData.Length; i++) data.Points.AddXY(i + 1, r.RawData[i]);
            chart.Series.Add(data);

            AddControlLines(chart, r);
            AddSpecificationLines(chart, r);

            chart.Titles.Add(new Title($"{r.ItemName} - 개별값 관리도")
            { Font = new SD.Font("Arial", 10, SD.FontStyle.Bold) });
            chart.Legends.Add(new Legend("Legend") { Docking = Docking.Right });
            return chart;
        }

        private static Dictionary<double, int> CreateHistogram(double[] data, int binCount)
        {
            double min = data.Min();
            double max = data.Max();
            double binWidth = (max - min) / binCount;
            if (binWidth <= 0) binWidth = 1e-9;

            var hist = new Dictionary<double, int>();
            for (int i = 0; i < binCount; i++)
            {
                double binStart = min + i * binWidth;
                hist[binStart + binWidth / 2] = 0;
            }

            foreach (double v in data)
            {
                int idx = Math.Min((int)((v - min) / binWidth), binCount - 1);
                double center = min + idx * binWidth + binWidth / 2;
                hist[center]++;
            }

            return hist;
        }

        private static void AddControlLines(Chart chart, ProcessCapabilityAnalyzer.CapabilityResult r)
        {
            var center = new Series("중심선")
            {
                ChartType = SeriesChartType.Line,
                Color = SD.Color.Green,
                BorderWidth = 2,
                BorderDashStyle = ChartDashStyle.Dash
            };
            center.Points.AddXY(1, r.Mean);
            center.Points.AddXY(r.RawData.Length, r.Mean);
            chart.Series.Add(center);

            double ucl = r.Mean + 3 * r.StdDev;
            double lcl = r.Mean - 3 * r.StdDev;

            var sU = new Series("상부관리한계")
            {
                ChartType = SeriesChartType.Line,
                Color = SD.Color.Red,
                BorderDashStyle = ChartDashStyle.Dash
            };
            sU.Points.AddXY(1, ucl);
            sU.Points.AddXY(r.RawData.Length, ucl);
            chart.Series.Add(sU);

            var sL = new Series("하부관리한계")
            {
                ChartType = SeriesChartType.Line,
                Color = SD.Color.Red,
                BorderDashStyle = ChartDashStyle.Dash
            };
            sL.Points.AddXY(1, lcl);
            sL.Points.AddXY(r.RawData.Length, lcl);
            chart.Series.Add(sL);
        }

        private static void AddSpecificationLines(Chart chart, ProcessCapabilityAnalyzer.CapabilityResult r)
        {
            if (r.LSL > double.MinValue)
            {
                var s = new Series("하한규격")
                {
                    ChartType = SeriesChartType.Line,
                    Color = SD.Color.Orange,
                    BorderWidth = 3,
                    BorderDashStyle = ChartDashStyle.DashDot
                };
                double maxY = chart.Series
                       .SelectMany(s => s.Points)
                       .Max(p => p.YValues[0]);
                s.Points.AddXY(r.LSL, 0); s.Points.AddXY(r.LSL, maxY);
                chart.Series.Add(s);
            }

            if (r.USL < double.MaxValue)
            {
                var s = new Series("상한규격")
                {
                    ChartType = SeriesChartType.Line,
                    Color = SD.Color.Orange,
                    BorderWidth = 3,
                    BorderDashStyle = ChartDashStyle.DashDot
                };
                double maxY = chart.Series
                       .SelectMany(s => s.Points)
                       .Max(p => p.YValues[0]);
                s.Points.AddXY(r.USL, 0); s.Points.AddXY(r.USL, maxY);
                chart.Series.Add(s);
            }
        }

        private static double NormalPDF(double x, double mean, double std)
        {
            double c = 1.0 / (std * Math.Sqrt(2 * Math.PI));
            double e = -0.5 * Math.Pow((x - mean) / std, 2);
            return c * Math.Exp(e);
        }
    }

    // ----------------------- 메인 폼 -----------------------
    public partial class CPKAnalysisForm : Form
    {
        private MenuStrip menuStrip = new MenuStrip();
        private ToolStripMenuItem fileMenu = new ToolStripMenuItem();
        private SplitContainer mainSplitContainer = new SplitContainer();
        private SplitContainer leftSplitContainer = new SplitContainer();
        private DataGridView dataGridView = new DataGridView();
        private Panel chartPanel = new Panel();
        private TextBox resultsTextBox = new TextBox();
        private string? loadedFilePath;

        private List<TestDataParser.TestItem> testItems = new List<TestDataParser.TestItem>();
        private ProcessCapabilityAnalyzer.CapabilityResult? currentResult;

        public CPKAnalysisForm() { InitializeComponent(); }

        private void InitializeComponent()
        {
            this.Text = "CPK 분석 도구 - 테스트 데이터 파일 분석";
            this.Size = new SD.Size(1400, 900);
            this.StartPosition = FormStartPosition.CenterScreen;

            CreateMenu();

            mainSplitContainer.Dock = DockStyle.Fill;
            mainSplitContainer.Orientation = WFOrient.Horizontal;
            mainSplitContainer.SplitterDistance = 40;

            leftSplitContainer.Dock = DockStyle.Fill;
            leftSplitContainer.Orientation = WFOrient.Vertical;
            leftSplitContainer.SplitterDistance = 700;

            CreateDataGridView();
            CreateResultsPanel();
            CreateChartPanel();

            leftSplitContainer.Panel1.Controls.Add(dataGridView);
            leftSplitContainer.Panel2.Controls.Add(resultsTextBox);

            mainSplitContainer.Panel1.Controls.Add(leftSplitContainer);
            mainSplitContainer.Panel2.Controls.Add(chartPanel);

            this.Controls.Add(mainSplitContainer);
            this.MainMenuStrip = menuStrip;
            this.Controls.Add(menuStrip);
        }

        private void CreateMenu()
        {
            menuStrip = new MenuStrip();
            fileMenu = new ToolStripMenuItem("파일(&F)");

            var openMenuItem = new ToolStripMenuItem("파일 열기(&O)");
            openMenuItem.Click += OpenFile_Click;
            openMenuItem.ShortcutKeys = Keys.Control | Keys.O;

            var exitMenuItem = new ToolStripMenuItem("종료(&X)");
            exitMenuItem.Click += (s, e) => this.Close();

            fileMenu.DropDownItems.Add(openMenuItem);
            fileMenu.DropDownItems.Add(new ToolStripSeparator());
            fileMenu.DropDownItems.Add(exitMenuItem);
            menuStrip.Items.Add(fileMenu);

            var analysisMenu = new ToolStripMenuItem("분석(&A)");
            var analyzeMenuItem = new ToolStripMenuItem("선택된 항목 분석(&A)");
            analyzeMenuItem.Click += AnalyzeSelected_Click;
            analyzeMenuItem.ShortcutKeys = Keys.F5;

            var exportMenuItem = new ToolStripMenuItem("결과 내보내기(&E)");
            exportMenuItem.Click += ExportResults_Click;

            var exportPdfMenuItem = new ToolStripMenuItem("PDF로 내보내기(&P)");
            exportPdfMenuItem.Click += ExportPdf_Click;

            analysisMenu.DropDownItems.Add(exportPdfMenuItem);
            analysisMenu.DropDownItems.Add(analyzeMenuItem);
            analysisMenu.DropDownItems.Add(exportMenuItem);
            menuStrip.Items.Add(analysisMenu);
        }

        private void CreateDataGridView()
        {
            dataGridView.Dock = DockStyle.Fill;
            dataGridView.AutoGenerateColumns = false;
            dataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView.MultiSelect = false;
            dataGridView.ReadOnly = true;
            dataGridView.AllowUserToAddRows = false;

            dataGridView.Columns.Add(new DataGridViewTextBoxColumn { Name = "Step", HeaderText = "Step", DataPropertyName = "Step", Width = 250 });
            dataGridView.Columns.Add(new DataGridViewTextBoxColumn { Name = "ItemName", HeaderText = "항목명", DataPropertyName = "ItemName", Width = 250 });
            dataGridView.Columns.Add(new DataGridViewTextBoxColumn { Name = "Spec", HeaderText = "규격", DataPropertyName = "Spec", Width = 100 });
            dataGridView.Columns.Add(new DataGridViewTextBoxColumn { Name = "SampleCount", HeaderText = "샘플수", DataPropertyName = "SampleCount", Width = 70 });
            dataGridView.Columns.Add(new DataGridViewTextBoxColumn { Name = "Mean", HeaderText = "평균", DataPropertyName = "Mean", Width = 80 });
            dataGridView.Columns.Add(new DataGridViewTextBoxColumn { Name = "StdDev", HeaderText = "표준편차", DataPropertyName = "StdDev", Width = 80 });
            dataGridView.Columns.Add(new DataGridViewTextBoxColumn { Name = "QuickEval", HeaderText = "빠른평가", DataPropertyName = "QuickEval", Width = 80 });
            dataGridView.Columns.Add(new DataGridViewTextBoxColumn { Name = "Unit", HeaderText = "단위", DataPropertyName = "Unit", Width = 50 });

            dataGridView.CellFormatting += DataGridView_CellFormatting;
            dataGridView.SelectionChanged += DataGridView_SelectionChanged;
        }

        private void DataGridView_CellFormatting(object? sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.ColumnIndex != dataGridView.Columns["QuickEval"].Index || e.Value == null) return;

            string evaluation = e.Value.ToString() ?? "";
            switch (evaluation)
            {
                case "≥ 2.0":
                    e.CellStyle.BackColor = SD.Color.LimeGreen; e.CellStyle.ForeColor = SD.Color.White;
                    e.CellStyle.Font = new SD.Font(e.CellStyle.Font, SD.FontStyle.Bold); break;
                case "≥ 1.67":
                    e.CellStyle.BackColor = SD.Color.Green; e.CellStyle.ForeColor = SD.Color.White;
                    e.CellStyle.Font = new SD.Font(e.CellStyle.Font, SD.FontStyle.Bold); break;
                case "≥ 1.33":
                    e.CellStyle.BackColor = SD.Color.YellowGreen; e.CellStyle.ForeColor = SD.Color.Black; break;
                case "≥ 1.0":
                    e.CellStyle.BackColor = SD.Color.Yellow; e.CellStyle.ForeColor = SD.Color.Black; break;
                case "≥ 0.67":
                    e.CellStyle.BackColor = SD.Color.Orange; e.CellStyle.ForeColor = SD.Color.Black; break;
                case "< 0.67":
                    e.CellStyle.BackColor = SD.Color.Red; e.CellStyle.ForeColor = SD.Color.White;
                    e.CellStyle.Font = new SD.Font(e.CellStyle.Font, SD.FontStyle.Bold); break;
                case "완벽":
                    e.CellStyle.BackColor = SD.Color.Gold; e.CellStyle.ForeColor = SD.Color.Black;
                    e.CellStyle.Font = new SD.Font(e.CellStyle.Font, SD.FontStyle.Bold); break;
                case "규격외":
                    e.CellStyle.BackColor = SD.Color.DarkRed; e.CellStyle.ForeColor = SD.Color.White;
                    e.CellStyle.Font = new SD.Font(e.CellStyle.Font, SD.FontStyle.Bold); break;
                case "데이터부족":
                case "계산오류":
                    e.CellStyle.BackColor = SD.Color.LightGray; e.CellStyle.ForeColor = SD.Color.Black; break;
            }
        }

        private void CreateResultsPanel()
        {
            resultsTextBox.Multiline = true;
            resultsTextBox.ScrollBars = ScrollBars.Vertical;
            resultsTextBox.Font = new SD.Font("Malgun Gothic", 9);
            resultsTextBox.Dock = DockStyle.Fill;
            resultsTextBox.ReadOnly = true;
            resultsTextBox.Text = "파일을 열어서 분석할 데이터를 로드하세요.";
        }

        private void CreateChartPanel()
        {
            chartPanel.AutoScroll = true;
            chartPanel.Dock = DockStyle.Fill;
            chartPanel.BackColor = SD.Color.White;
        }

        private void OpenFile_Click(object? sender, EventArgs e)
        {
            using var ofd = new OpenFileDialog
            {
                Filter = "csv 파일 (*.csv)|*.csv|모든 파일 (*.*)|*.*",
                FilterIndex = 1
            };
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                try { LoadTestData(ofd.FileName); }
                catch (Exception ex)
                {
                    MessageBox.Show($"파일 로드 중 오류: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void LoadTestData(string filePath)
        {
            loadedFilePath = filePath;
            testItems = TestDataParser.ParseFile(filePath);
            if (!testItems.Any())
            {
                MessageBox.Show("유효한 테스트 데이터를 찾을 수 없습니다.", "경고", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var display = testItems.Select(item => new
            {
                Step = item.Step,
                ItemName = item.ItemName,
                Spec = item.OriginalSpec,
                SampleCount = item.Values.Count(v => v != 0),
                Mean = item.Values.Where(v => v != 0).Any()
                        ? item.Values.Where(v => v != 0).Average().ToString("F4") : "N/A",
                StdDev = item.Values.Where(v => v != 0).Count() > 1
                        ? CalculateStdDev(item.Values.Where(v => v != 0)).ToString("F4") : "N/A",
                QuickEval = GetQuickEvaluation(item),
                Unit = item.Unit
            }).ToList();

            dataGridView.DataSource = display;

            resultsTextBox.Text =
                $"총 {testItems.Count}개의 테스트 항목을 로드했습니다.\n" +
                "분석할 항목을 선택하고 F5를 누르거나 '분석 > 선택된 항목 분석'을 클릭하세요.\n\n" +
                "빠른평가 색상 범례 (Cpk 값 기준):\n" +
                "■ 초록(진한) - ≥ 2.0\n" +
                "■ 초록(연한) - ≥ 1.67\n" +
                "■ 연두 - ≥ 1.33\n" +
                "■ 노랑 - ≥ 1.0\n" +
                "■ 주황 - ≥ 0.67\n" +
                "■ 빨강 - < 0.67\n" +
                "■ 금색 - 완벽\n" +
                "■ 암적색 - 규격외";
        }

        private double CalculateStdDev(IEnumerable<double> values)
        {
            var arr = values.ToArray();
            if (arr.Length <= 1) return 0;
            double mean = arr.Average();
            double ss = arr.Sum(x => Math.Pow(x - mean, 2));
            double v = Math.Sqrt(ss / (arr.Length - 1));
            return v < 1e-10 ? 0 : v;
        }

        private static double CalculateC4(int n)
        {
            if (n < 2) return 1.0;
            var c4Table = new Dictionary<int, double> {
                {2, 0.7979},{3,0.8862},{4,0.9213},{5,0.94},
                {6,0.9515},{7,0.9594},{8,0.965},{9,0.9693},
                {10,0.9727},{15,0.9823},{20,0.9869},{25,0.9896}
            };
            if (c4Table.ContainsKey(n)) return c4Table[n];
            if (n > 25) return 1.0 - 1.0 / (4 * n) + 7.0 / (32 * n * n);

            var keys = c4Table.Keys.OrderBy(x => x).ToArray();
            for (int i = 0; i < keys.Length - 1; i++)
                if (n > keys[i] && n < keys[i + 1])
                {
                    double t = (double)(n - keys[i]) / (keys[i + 1] - keys[i]);
                    return c4Table[keys[i]] + t * (c4Table[keys[i + 1]] - c4Table[keys[i]]);
                }
            return 0.97;
        }

        private string GetQuickEvaluation(TestDataParser.TestItem item)
        {
            var valid = item.Values.Where(v => v != 0).ToArray();
            if (valid.Length < 2) return "데이터부족";

            try
            {
                double mean = valid.Average();
                double overallStd = CalculateStdDev(valid);

                if (overallStd == 0)
                {
                    bool within = mean >= item.LSL && mean <= item.USL;
                    return within ? "완벽" : "규격외";
                }

                int n = valid.Length;
                double c4 = CalculateC4(n);
                double withinStd = overallStd / c4;

                double cpl = (mean - item.LSL) / (3 * withinStd);
                double cpu = (item.USL - mean) / (3 * withinStd);
                double cpk = Math.Min(cpl, cpu);

                if (double.IsInfinity(cpk) || double.IsNaN(cpk)) return "계산오류";
                if (cpk >= 2.0) return "≥ 2.0";
                else if (cpk >= 1.67) return "≥ 1.67";
                else if (cpk >= 1.33) return "≥ 1.33";
                else if (cpk >= 1.0) return "≥ 1.0";
                else if (cpk >= 0.67) return "≥ 0.67";
                else return "< 0.67";
            }
            catch { return "계산오류"; }
        }

        private void DataGridView_SelectionChanged(object? sender, EventArgs e)
        {
            if (dataGridView.SelectedRows.Count == 0 || testItems.Count == 0) return;
            int idx = dataGridView.SelectedRows[0].Index;
            if (idx < 0 || idx >= testItems.Count) return;
            UpdatePreview(testItems[idx]);
        }

        private void UpdatePreview(TestDataParser.TestItem item)
        {
            var valid = item.Values.Where(v => v != 0).ToArray();
            if (!valid.Any()) { resultsTextBox.Text = "선택된 항목에 유효한 데이터가 없습니다."; return; }

            string quick = GetQuickEvaluation(item);

            resultsTextBox.Text =
                $"선택된 항목: {item.ItemName}\n" +
                $"규격: {item.OriginalSpec} {item.Unit}\n" +
                $"유효 샘플 수: {valid.Length}\n" +
                $"평균: {valid.Average():F4}\n" +
                $"표준편차: {CalculateStdDev(valid):F4}\n" +
                $"최소값: {valid.Min():F4}\n" +
                $"최대값: {valid.Max():F4}\n" +
                $"평가: {quick}\n\n" +
                "F5를 눌러 상세 CPK 분석을 수행하세요.";
        }

        private void AnalyzeSelected_Click(object? sender, EventArgs e)
        {
            if (dataGridView.SelectedRows.Count == 0 || testItems.Count == 0)
            {
                MessageBox.Show("분석할 항목을 선택하세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            int idx = dataGridView.SelectedRows[0].Index;
            var item = testItems[idx];

            var valid = item.Values.Where(v => v != 0).ToArray();
            if (valid.Length < 2)
            {
                MessageBox.Show("CPK 분석을 위해서는 최소 2개 이상의 유효한 데이터가 필요합니다.",
                    "경고", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                currentResult = ProcessCapabilityAnalyzer.CalculateProcessCapability(
                    valid, item.LSL, item.USL, item.ItemName, item.Unit, true, item.Step);

                UpdateResults();
                UpdateCharts();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"분석 중 오류: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void UpdateResults()
        {
            if (currentResult == null) return;
            var r = currentResult;

            resultsTextBox.Text = $@"=== CPK 분석 결과 ===
스텝명: {r.Step}     
항목명: {r.ItemName}
단위: {r.Unit}
샘플 수: {r.SampleSize}
평균: {r.Mean:F4}
표준편차: {r.StdDev:F4}
하한규격 (LSL): {r.LSL:F4}
상한규격 (USL): {r.USL:F4}

=== 공정능력지수 ===
Cp:  {r.Cp:F4}
Cpk: {r.Cpk:F4}
  - Cpk(Lower): {r.CpkLower:F4}
  - Cpk(Upper): {r.CpkUpper:F4}
Pp:  {r.Pp:F4}
Ppk: {r.Ppk:F4}

=== 공정능력 평가 ===
- Cpk ≥ 1.33: 공정능력 양호
- 1.0 ≤ Cpk < 1.33: 공정능력 보통 (개선 권장)
- Cpk < 1.0: 공정능력 불량 (즉시 개선 필요)

현재 상태: {(r.IsCapable ? "양호 (Cpk ≥ 1.33)" : r.Cpk >= 1.0 ? "보통 (개선 권장)" : "불량 (즉시 개선 필요)")}

=== 불량률 추정 ===
예상 불량률: {EstimateDefectRate(r):F2} PPM
- LSL 초과 불량률: {EstimateDefectRateLSL(r):F2} PPM
- USL 초과 불량률: {EstimateDefectRateUSL(r):F2} PPM";
        }

        private double EstimateDefectRate(ProcessCapabilityAnalyzer.CapabilityResult r)
        {
            double zLSL = (r.Mean - r.LSL) / r.StdDev;
            double zUSL = (r.USL - r.Mean) / r.StdDev;
            double defectLSL = NormalCDF(-zLSL) * 1_000_000;
            double defectUSL = NormalCDF(-zUSL) * 1_000_000;
            return defectLSL + defectUSL;
        }
        private double EstimateDefectRateLSL(ProcessCapabilityAnalyzer.CapabilityResult r)
            => NormalCDF(-(r.Mean - r.LSL) / r.StdDev) * 1_000_000;
        private double EstimateDefectRateUSL(ProcessCapabilityAnalyzer.CapabilityResult r)
            => NormalCDF(-(r.USL - r.Mean) / r.StdDev) * 1_000_000;

        private double NormalCDF(double x) => 0.5 * (1.0 + ErrorFunction(x / Math.Sqrt(2.0)));
        private double ErrorFunction(double x)
        {
            double a1 = 0.254829592, a2 = -0.284496736, a3 = 1.421413741,
                   a4 = -1.453152027, a5 = 1.061405429, p = 0.3275911;
            int sign = x < 0 ? -1 : 1;
            x = Math.Abs(x);
            double t = 1.0 / (1.0 + p * x);
            double y = 1.0 - (((((a5 * t + a4) * t) + a3) * t + a2) * t + a1) * t * Math.Exp(-x * x);
            return sign * y;
        }

        private void UpdateCharts()
        {
            if (currentResult == null) return;
            chartPanel.Controls.Clear();

            try
            {
                if (currentResult.StdDev == 0)
                {
                    var infoPanel = CreateZeroVariationPanel();
                    infoPanel.Location = new SD.Point(10, 10);
                    chartPanel.Controls.Add(infoPanel);
                }
                else
                {
                    // ✅ 히스토그램만
                    var hist = CPKChartGenerator.CreateCapabilityHistogram(currentResult);
                    hist.Location = new SD.Point(10, 10);
                    chartPanel.Controls.Add(hist);

                    // ⛔ 개별값 관리도 제거
                    // var ctrl = CPKChartGenerator.CreateControlChart(currentResult);
                    // ctrl.Location = new SD.Point(10, 530);
                    // chartPanel.Controls.Add(ctrl);

                    var summary = CreateSummaryPanel();
                    summary.Location = new SD.Point(830, 10);
                    chartPanel.Controls.Add(summary);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"차트 생성 중 오류: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string BuildSummaryText(ProcessCapabilityAnalyzer.CapabilityResult r)
        {
            return
                $@"Step: {r.Step}
        항목: {r.ItemName}
단위: {r.Unit}
샘플수: {r.SampleSize}

규격 정보:
LSL: {r.LSL:F4}
USL: {r.USL:F4}
규격폭: {(r.USL - r.LSL):F4}

통계량:
평균: {r.Mean:F4}
표준편차: {r.StdDev:F4}
범위: {(r.RawData.Max() - r.RawData.Min()):F4}

공정능력지수:
Cp = {r.Cp:F4}
Cpk = {r.Cpk:F4}
Pp = {r.Pp:F4}
Ppk = {r.Ppk:F4}

평가:
{(r.IsCapable ? "공정능력 양호" : r.Cpk >= 1.0 ? "개선 권장" : "즉시 개선 필요")}

예상 불량률:
{EstimateDefectRate(r):F1} PPM

중심화:
목표중심: {((r.LSL + r.USL) / 2):F4}
현재중심: {r.Mean:F4}
편차: {(r.Mean - (r.LSL + r.USL) / 2):F4}";
        }


        private Panel CreateZeroVariationPanel()
        {
            var panel = new Panel
            {
                Size = new SD.Size(800, 400),
                BorderStyle = WFBorderStyle.FixedSingle,
                BackColor = SD.Color.LightYellow
            };

            var title = new Label
            {
                Text = "변동 없음 (모든 데이터 동일)",
                Font = new SD.Font("Arial", 16, SD.FontStyle.Bold),
                Location = new SD.Point(20, 20),
                Size = new SD.Size(760, 30),
                TextAlign = SD.ContentAlignment.MiddleCenter,
                ForeColor = SD.Color.DarkBlue
            };
            panel.Controls.Add(title);

            var info = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Location = new SD.Point(20, 70),
                Size = new SD.Size(760, 310),
                Font = new SD.Font("Malgun Gothic", 12),
                BackColor = SD.Color.White
            };

            if (currentResult != null)
            {
                bool within = currentResult.Mean >= currentResult.LSL && currentResult.Mean <= currentResult.USL;
                info.Text = $@"분석 결과: {currentResult.ItemName}

상황 분석:
   모든 측정값이 {currentResult.Mean:F4} {currentResult.Unit}로 동일합니다.
   이는 측정 시스템의 분해능 부족이거나 실제로 변동이 없는 경우입니다.

규격 정보:
   하한규격 (LSL): {currentResult.LSL:F4} {currentResult.Unit}
   상한규격 (USL): {currentResult.USL:F4} {currentResult.Unit}
   측정값: {currentResult.Mean:F4} {currentResult.Unit}

공정능력 평가:
   {(within ? "완벽한 상태" : "규격 외 상태")}
   {(within ? "  - 모든 값이 규격 내에 있음" : "  - 모든 값이 규격을 벗어남")}
   {(within ? "  - 변동이 없어 이상적인 상태" : "  - 즉시 조치가 필요함")}

권장 사항:
   {(within ?
     "• 현재 상태를 유지하세요\n   • 측정 분해능 향상을 고려해보세요\n   • 정기적인 모니터링을 계속하세요" :
     "• 즉시 공정을 점검하세요\n   • 원인을 파악하여 조치하세요\n   • 규격 내로 조정이 필요합니다")}

참고:
   일반적인 CPK 계산은 변동이 있어야 의미가 있습니다.
   현재 상태에서는 통계적 관리보다는 평균값 관리가 중요합니다.";
            }

            panel.Controls.Add(info);
            return panel;
        }

        private Panel CreateSummaryPanel()
        {
            var panel = new Panel
            {
                Size = new SD.Size(300, 470),
                BorderStyle = WFBorderStyle.FixedSingle,
                BackColor = SD.Color.White
            };

            var title = new Label
            {
                Text = "CPK 분석 요약",
                Font = new SD.Font("Arial", 12, SD.FontStyle.Bold),
                Location = new SD.Point(10, 10),
                Size = new SD.Size(280, 25),
                TextAlign = SD.ContentAlignment.MiddleCenter
            };
            panel.Controls.Add(title);

            var tb = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Location = new SD.Point(10, 45),
                Size = new SD.Size(280, 410),
                Font = new SD.Font("Malgun Gothic", 8)
            };

            if (currentResult != null)
            {
                tb.Text = $@"항목: {currentResult.ItemName}
단위: {currentResult.Unit}
샘플수: {currentResult.SampleSize}

규격 정보:
LSL: {currentResult.LSL:F4}
USL: {currentResult.USL:F4}
규격폭: {(currentResult.USL - currentResult.LSL):F4}

통계량:
평균: {currentResult.Mean:F4}
표준편차: {currentResult.StdDev:F4}
범위: {(currentResult.RawData.Max() - currentResult.RawData.Min()):F4}

공정능력지수:
Cp = {currentResult.Cp:F4}
Cpk = {currentResult.Cpk:F4}
Pp = {currentResult.Pp:F4}
Ppk = {currentResult.Ppk:F4}

평가:
{(currentResult.IsCapable ? "공정능력 양호" : currentResult.Cpk >= 1.0 ? "개선 권장" : "즉시 개선 필요")}

예상 불량률:
{EstimateDefectRate(currentResult):F1} PPM

중심화:
목표중심: {((currentResult.LSL + currentResult.USL) / 2):F4}
현재중심: {currentResult.Mean:F4}
편차: {(currentResult.Mean - (currentResult.LSL + currentResult.USL) / 2):F4}";
            }

            panel.Controls.Add(tb);
            return panel;
        }

        private void ExportResults_Click(object? sender, EventArgs e)
        {
            if (currentResult == null)
            {
                MessageBox.Show("먼저 분석을 수행하세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using var sfd = new SaveFileDialog
            {
                Filter = "CSV 파일 (*.csv)|*.csv|텍스트 파일 (*.txt)|*.txt",
                FilterIndex = 1,
                FileName = $"CPK_Analysis_{currentResult.ItemName.Replace(" ", "_")}_{DateTime.Now:yyyyMMdd_HHmmss}"
            };
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try { ExportToFile(sfd.FileName); MessageBox.Show("결과가 성공적으로 내보내졌습니다.", "완료", MessageBoxButtons.OK, MessageBoxIcon.Information); }
                catch (Exception ex) { MessageBox.Show($"파일 내보내기 중 오류: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            }
        }

        private void ExportToFile(string filePath)
        {
            if (currentResult == null) return;
            var r = currentResult;
            var lines = new List<string>();

            if (filePath.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
            {
                lines.Add("항목,값");
                lines.Add($"분석일시,{DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                lines.Add($"항목명,{r.ItemName}");
                lines.Add($"단위,{r.Unit}");
                lines.Add($"샘플수,{r.SampleSize}");
                lines.Add($"평균,{r.Mean:F6}");
                lines.Add($"표준편차,{r.StdDev:F6}");
                lines.Add($"하한규격,{r.LSL:F6}");
                lines.Add($"상한규격,{r.USL:F6}");
                lines.Add($"Cp,{r.Cp:F6}");
                lines.Add($"Cpk,{r.Cpk:F6}");
                lines.Add($"CpkLower,{r.CpkLower:F6}");
                lines.Add($"CpkUpper,{r.CpkUpper:F6}");
                lines.Add($"Pp,{r.Pp:F6}");
                lines.Add($"Ppk,{r.Ppk:F6}");
                lines.Add($"공정능력평가,{(r.IsCapable ? "양호" : r.Cpk >= 1.0 ? "보통" : "불량")}");
                lines.Add($"예상불량률(PPM),{EstimateDefectRate(r):F2}");
                lines.Add("");
                lines.Add("원시데이터");
                lines.Add("순번,측정값");
                for (int i = 0; i < r.RawData.Length; i++)
                    lines.Add($"{i + 1},{r.RawData[i]:F6}");
            }
            else
            {
                lines.Add("=== CPK 공정능력 분석 결과 ===");
                lines.Add($"분석일시: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                lines.Add($"항목명: {r.ItemName}");
                lines.Add($"단위: {r.Unit}");
                lines.Add("");
                lines.Add("=== 기본 통계량 ===");
                lines.Add($"샘플 수: {r.SampleSize}");
                lines.Add($"평균: {r.Mean:F6}");
                lines.Add($"표준편차: {r.StdDev:F6}");
                lines.Add($"최소값: {r.RawData.Min():F6}");
                lines.Add($"최대값: {r.RawData.Max():F6}");
                lines.Add($"범위: {(r.RawData.Max() - r.RawData.Min()):F6}");
                lines.Add("");
                lines.Add("=== 규격 정보 ===");
                lines.Add($"하한규격 (LSL): {r.LSL:F6}");
                lines.Add($"상한규격 (USL): {r.USL:F6}");
                lines.Add($"규격 폭: {(r.USL - r.LSL):F6}");
                lines.Add("");
                lines.Add("=== 공정능력지수 ===");
                lines.Add($"Cp:  {r.Cp:F6}");
                lines.Add($"Cpk: {r.Cpk:F6}");
                lines.Add($"  - CpkLower: {r.CpkLower:F6}");
                lines.Add($"  - CpkUpper: {r.CpkUpper:F6}");
                lines.Add($"Pp:  {r.Pp:F6}");
                lines.Add($"Ppk: {r.Ppk:F6}");
                lines.Add("");
                lines.Add("=== 공정능력 평가 ===");
                lines.Add($"평가 결과: {(r.IsCapable ? "양호 (Cpk ≥ 1.33)" : r.Cpk >= 1.0 ? "보통 (개선 권장)" : "불량 (즉시 개선 필요)")}");
                lines.Add($"예상 불량률: {EstimateDefectRate(r):F2} PPM");
                lines.Add($"  - LSL 초과: {EstimateDefectRateLSL(r):F2} PPM");
                lines.Add($"  - USL 초과: {EstimateDefectRateUSL(r):F2} PPM");
                lines.Add("");
                lines.Add("=== 원시 데이터 ===");
                for (int i = 0; i < r.RawData.Length; i++)
                    lines.Add($"{i + 1:D3}: {r.RawData[i]:F6}");
            }

            File.WriteAllLines(filePath, lines, System.Text.Encoding.UTF8);
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.F5) { AnalyzeSelected_Click(null, null); return true; }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        // ----------------------- PDF 내보내기 (문자열 경로 AddImage) -----------------------
        private void ExportPdf_Click(object? sender, EventArgs e)
        {
            if (testItems.Count == 0)
            {
                MessageBox.Show("먼저 파일을 로드하세요.", "알림",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 불러온 파일명 기반 기본 이름 만들기
            string baseName = "CPK_Report";
            if (!string.IsNullOrWhiteSpace(loadedFilePath))
                baseName = Path.GetFileNameWithoutExtension(loadedFilePath);

            // 파일명에서 금지문자 제거 (안전)
            char[] bad = Path.GetInvalidFileNameChars();
            baseName = string.Concat(baseName.Select(c => bad.Contains(c) ? '_' : c));

            string stamp = DateTime.Now.ToString("yyyyMMdd");
            string defaultFileName = $"{baseName}_CPK_{stamp}.pdf";

            using var sfd = new SaveFileDialog
            {
                Filter = "PDF 파일 (*.pdf)|*.pdf",
                AddExtension = true,
                DefaultExt = "pdf",
                FileName = defaultFileName,
                // 탐색기 시작 위치를 원본 파일 폴더로
                InitialDirectory = !string.IsNullOrWhiteSpace(loadedFilePath)
                    ? Path.GetDirectoryName(loadedFilePath)
                    : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                OverwritePrompt = true
            };

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    ExportToPdf(sfd.FileName);
                    MessageBox.Show("PDF로 내보냈습니다.", "완료",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"PDF 내보내기 중 오류: {ex.Message}", "오류",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }


        private void ExportToPdf(string filePath)
        {
            string SaveChartImage(Chart ch)
            {
                string tmp = Path.Combine(Path.GetTempPath(), $"cpk_chart_{Guid.NewGuid():N}.png");
                ch.AntiAliasing = AntiAliasingStyles.All;
                ch.TextAntiAliasingQuality = TextAntiAliasingQuality.High;
                foreach (ChartArea a in ch.ChartAreas) a.RecalculateAxesScale();
                ch.Invalidate();
                ch.SaveImage(tmp, ChartImageFormat.Png);
                return tmp;
            }


            void InsertChartsForResult(Section section, ProcessCapabilityAnalyzer.CapabilityResult r, List<string> tempFiles)
            {
                if (r.StdDev <= 1e-10)
                {
                    // 변동 0 – 간단한 설명 문단으로 대체
                    var p = section.AddParagraph();
                    p.Format.Font.Name = "Malgun Gothic";
                    p.Format.SpaceAfter = "3mm";
                    p.AddFormattedText($"{r.ItemName} - 변동 없음(모든 값이 동일)", TextFormat.Bold);
                    p.AddLineBreak();
                    p.AddText($"측정값 {r.Mean:F4} {r.Unit}, LSL {r.LSL:F4}, USL {r.USL:F4} – 차트 생략");
                    return;
                }

                // 히스토그램
                var hist = CPKChartGenerator.CreateCapabilityHistogram(r);
                string histPath = SaveChartImage(hist);
                tempFiles.Add(histPath);
                AddImageFlexible(section, histPath, 24.0);
                section.AddParagraph().Format.SpaceAfter = "3mm";

                // 관리도
                var ctrl = CPKChartGenerator.CreateControlChart(r);
                string ctrlPath = SaveChartImage(ctrl);
                tempFiles.Add(ctrlPath);
                AddImageFlexible(section, ctrlPath, 24.0);
                section.AddParagraph().Format.SpaceAfter = "3mm";

                // 메모리 누수 방지
                hist.Dispose();
                ctrl.Dispose();
            }

            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("저장 경로가 비어 있습니다.", nameof(filePath));

            if (this.dataGridView == null)
                throw new NullReferenceException("dataGridView가 초기화되지 않았습니다.");
            if (this.resultsTextBox == null)
                throw new NullReferenceException("resultsTextBox가 초기화되지 않았습니다.");
            if (this.chartPanel == null)
                throw new NullReferenceException("chartPanel이 초기화되지 않았습니다.");

            // 1) 문서 & 스타일
            var doc = new Document();
            doc.Info.Title = "CPK 분석 보고서";

            // 기본 글꼴: 맑은 고딕
            var normal = doc.Styles["Normal"];
            normal.Font.Name = "Malgun Gothic";
            normal.Font.Size = 10;

            var heading = doc.Styles.AddStyle("Heading", "Normal");
            heading.Font.Size = 14;
            heading.Font.Bold = true;

            // 2) 섹션 & 용지 세팅 (A4 가로)
            var sec = doc.AddSection();
            sec.PageSetup.PageFormat = PageFormat.A4;
            sec.PageSetup.Orientation = MigraDocCore.DocumentObjectModel.Orientation.Landscape;
            sec.PageSetup.TopMargin = "15mm";
            sec.PageSetup.BottomMargin = "15mm";
            sec.PageSetup.LeftMargin = "15mm";
            sec.PageSetup.RightMargin = "15mm";

            // 3) 표지/개요
            var title = sec.AddParagraph("CPK 분석 보고서", "Heading");
            title.Format.SpaceAfter = "6mm";

            var info = sec.AddParagraph();
            info.AddText($"생성일시: {DateTime.Now:yyyy-MM-dd HH:mm:ss}\n");
            info.AddText($"총 항목 수: {this.testItems?.Count ?? 0}\n");
            info.Format.SpaceAfter = "6mm";

            // 4) 상세 결과(우측 텍스트박스)
            /*var resTitle = sec.AddParagraph("상세 결과 (선택 항목):", "Heading");
            resTitle.Format.SpaceBefore = "3mm";

            var res = sec.AddParagraph();
            res.Format.Font.Name = "Malgun Gothic";   // 한글 깨짐 방지 (모노스페이스 사용 X)
            res.Format.Font.Size = 9;
            res.AddText(this.resultsTextBox.Text);
            res.Format.SpaceAfter = "8mm";*/

            // 5) 데이터 요약 테이블 (DataGridView)
            var tblTitle = sec.AddParagraph("데이터 요약(그리드):", "Heading");
            tblTitle.Format.SpaceBefore = "3mm";

            var table = sec.AddTable();
            table.Borders.Width = 0.5;
            table.Rows.LeftIndent = 0;
            table.Format.Font.Name = "Malgun Gothic";
            table.Format.Font.Size = 9;

            // 컬럼 생성 (간단 폭 규칙)
            foreach (DataGridViewColumn col in this.dataGridView.Columns)
            {
                double cm =
                    col.Width <= 60 ? 2.0 :
                    col.Width <= 90 ? 2.8 :
                    col.Width <= 120 ? 3.5 : 4.0;

                var c = table.AddColumn(Unit.FromCentimeter(cm));
                c.Format.Alignment = ParagraphAlignment.Center;
            }

            // 헤더 행
            var header = table.AddRow();
            header.HeadingFormat = true;
            header.Shading.Color = Colors.LightGray;
            header.Format.Font.Bold = true;

            for (int i = 0; i < this.dataGridView.Columns.Count; i++)
                header.Cells[i].AddParagraph(this.dataGridView.Columns[i].HeaderText ?? "");

            // 데이터 행
            foreach (DataGridViewRow dgvr in this.dataGridView.Rows)
            {
                if (dgvr.IsNewRow) continue;
                var r = table.AddRow();
                for (int i = 0; i < this.dataGridView.Columns.Count; i++)
                {
                    var val = dgvr.Cells[i].Value?.ToString() ?? "";
                    r.Cells[i].AddParagraph(val);
                }
            }

            var tempFilesList = new List<string>();
            try
            {
                // 섹션을 "차트 섹션"에서 시작하고 싶으면:
                var chartSecTitle = sec.AddParagraph("차트 및 분석 요약(가로 배치)", "Heading");
                chartSecTitle.Format.PageBreakBefore = true;   // ← 새 페이지에서 시작
                chartSecTitle.Format.SpaceAfter = "2mm";

                // 1개만
                if (currentResult != null)
                {
                    InsertChartAndSummarySideBySide(sec, currentResult, tempFilesList);
                }
                else
                {
                    // 또는 모든 항목 루프
                    foreach (var item in testItems)
                    {
                        var valid = item.Values.Where(v => v != 0).ToArray();
                        if (valid.Length < 2) continue;

                        var rEach = ProcessCapabilityAnalyzer.CalculateProcessCapability(
                                        valid, item.LSL, item.USL, item.ItemName, item.Unit, true, item.Step);

                        InsertChartAndSummarySideBySide(sec, rEach, tempFilesList);
                        // (항목마다 새 페이지를 쓰고 있다면 이 아래에 아무 것도 추가하지 마세요)
                    }

                }
            }
            finally
            {
                foreach (var f in tempFilesList)
                    try { if (File.Exists(f)) File.Delete(f); } catch { }
            }

            // 9) PDF 렌더 (유니코드/임베딩 권장)
            var renderer = new PdfDocumentRenderer(unicode: true)
            {
                Document = doc
            };
            renderer.RenderDocument();
            renderer.PdfDocument.Save(filePath);

            // ====== 로컬 헬퍼들 ======

            // 차트들을 PNG로 저장하고 섹션에 삽입
            string[] SaveChartsAndInsert(Section section, Panel panel)
            {
                if (section == null) throw new ArgumentNullException(nameof(section));
                if (panel == null) throw new ArgumentNullException(nameof(panel));

                var list = new System.Collections.Generic.List<string>();
                int idx = 1;

                foreach (Control ctrl in panel.Controls)
                {
                    if (ctrl is Chart ch)
                    {
                        // 저장
                        string tmp = Path.Combine(Path.GetTempPath(), $"cpk_chart_{idx++}_{Guid.NewGuid():N}.png");

                        // 렌더 품질 살짝 올리기
                        ch.AntiAliasing = AntiAliasingStyles.All;
                        ch.TextAntiAliasingQuality = TextAntiAliasingQuality.High;

                        ch.SaveImage(tmp, ChartImageFormat.Png);
                        list.Add(tmp);

                        // PDF에 삽입 (버전 자동대응)
                        AddImageFlexible(section, tmp, 24.0 /* A4 가로의 넓은 폭 확보 */);

                        var spacer = section.AddParagraph();
                        spacer.Format.SpaceAfter = "3mm";
                    }
                }

                return list.ToArray();
            }
            void InsertChartAndSummarySideBySide(
    Section section,
    ProcessCapabilityAnalyzer.CapabilityResult r,
    List<string> tmpFiles)
            {
                // 제목: 다음 표와 같은 페이지에 붙이기
                var h = section.AddParagraph($"{r.Step} - {r.ItemName} – 차트 & 분석 요약", "Heading");
                h.Format.SpaceBefore = "2mm";
                h.Format.SpaceAfter = "2mm";
                h.Format.KeepWithNext = true;          // ★ 제목 고아 방지 핵심

                // 2열 표 (왼쪽 차트, 오른쪽 요약)
                var table = section.AddTable();
                table.Borders.Width = 0;
                table.Rows.LeftIndent = 0;
                table.Format.Font.Name = "Malgun Gothic";
                table.Format.Font.Size = 9;

                var colChart = table.AddColumn(Unit.FromCentimeter(19.0));
                var colSummary = table.AddColumn(Unit.FromCentimeter(7.0));

                var row = table.AddRow();
                row.TopPadding = Unit.FromMillimeter(1.5);
                row.BottomPadding = Unit.FromMillimeter(1.5);

                // 왼쪽 셀: 차트 (또는 변동 0 안내)
                if (r.StdDev <= 1e-10)
                {
                    var p = row.Cells[0].AddParagraph(
                        $"변동 없음(모든 값 동일)\n" +
                        $"측정값 {r.Mean:F4} {r.Unit}\nLSL {r.LSL:F4}, USL {r.USL:F4}\n\n차트 생략"
                    );
                    p.Format.Font.Size = 10;
                }
                else
                {
                    var hist = CPKChartGenerator.CreateCapabilityHistogram(r);
                    string imgPath = SaveChartImage(hist);
                    tmpFiles.Add(imgPath);

                    var paraImg = row.Cells[0].AddParagraph();
                    var src = ImageSource.FromFile(imgPath);   // IImageSource
                    var img = paraImg.AddImage(src);           // 문자열 오버로드 X → IImageSource 사용
                    img.LockAspectRatio = true;
                    img.Width = Unit.FromCentimeter(18.5);     // 셀 폭 약간 여유

                    hist.Dispose();
                }

                // 오른쪽 셀: 요약 텍스트
                var summary = BuildSummaryText(r);
                var paraSum = row.Cells[1].AddParagraph(summary);
                paraSum.Format.Font.Size = 9;

                // 각 항목을 "항상" 다음 페이지에서 시작하고 싶다면 아래 추가:
                var pageBreak = section.AddParagraph();
                pageBreak.Format.PageBreakBefore = true;       // ★ 다음 항목은 새 페이지에서 시작
            }


            // MigraDocCore 버전별 AddImage(...) 자동 대응
            void AddImageFlexible(Section section, string imagePath, double widthCm)
            {
                if (!File.Exists(imagePath))
                    throw new FileNotFoundException($"이미지 파일이 없습니다: {imagePath}");

                // 여러분 환경은 IImageSource 오버로드가 열려있으니 이 경로 추천
                var src = ImageSource.FromFile(imagePath);
                var img = section.AddImage(src);
                img.LockAspectRatio = true;
                img.Width = Unit.FromCentimeter(widthCm);

                // 만약 위 줄이 컴파일 안되면, 아래 두 줄로 대체:
                // var img2 = section.AddImage(imagePath);
                // img2.LockAspectRatio = true; img2.Width = Unit.FromCentimeter(widthCm);
            }

            void InsertSummaryAndChartAsTwoPages(
    Section section,
    ProcessCapabilityAnalyzer.CapabilityResult r,
    List<string> tmpFiles)
            {
                // --- Page 1: 분석 요약 전용 페이지 ---
                // 제목
                var h1 = section.AddParagraph($"{r.ItemName} - 분석 요약", "Heading");
                h1.Format.SpaceBefore = "0mm";
                // 요약 본문
                var summaryPara = section.AddParagraph(BuildSummaryText(r));
                summaryPara.Format.Font.Name = "Malgun Gothic";
                summaryPara.Format.Font.Size = 9;

                // 요약 페이지 끝 → 다음 페이지로
                var pb1 = section.AddParagraph();
                pb1.Format.PageBreakBefore = true; // ★ 다음 페이지로 넘김

                // --- Page 2: 차트 전용 페이지 ---
                if (r.StdDev <= 1e-10)
                {
                    // 변동 0이면 차트 대신 안내만 (차트 페이지에도 안내)
                    var noVar = section.AddParagraph($"{r.ItemName} - 변동 없음(모든 값 동일)", "Heading");
                    noVar.Format.SpaceBefore = "0mm";

                    var txt = section.AddParagraph(
                        $"측정값 {r.Mean:F4} {r.Unit}, LSL {r.LSL:F4}, USL {r.USL:F4}\n" +
                        "표준편차가 0으로 차트를 생략합니다.");
                    txt.Format.Font.Name = "Malgun Gothic";
                    txt.Format.Font.Size = 9;
                }
                else
                {
                    // 히스토그램 즉석 생성 → PNG 저장 → 한 페이지 꽉 채우기
                    var hist = CPKChartGenerator.CreateCapabilityHistogram(r);
                    string histPath = SaveChartImage(hist);  // 기존 helper 사용
                    tmpFiles.Add(histPath);

                    // 이미지 삽입 (A4 가로, 좌우 15mm 여백 고려 — 24cm면 꽉 찹니다)
                    AddImageFlexible(section, histPath, 24.0); // 기존 helper 사용

                    hist.Dispose();
                }

                // 차트 페이지도 종료하고 다음 항목은 새 페이지에서 시작하게 설정
                var pb2 = section.AddParagraph();
                pb2.Format.PageBreakBefore = true; // ★ 다음 항목 대비 페이지 분리
            }

        }
    }
}
