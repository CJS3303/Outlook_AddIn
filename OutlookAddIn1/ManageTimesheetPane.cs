using System;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Threading.Tasks;
using Microsoft.Data.SqlClient;
using System.Data;
using System.Configuration;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn1
{
    [System.ComponentModel.DesignerCategory("Code")]
    public class ManageTimesheetPane : UserControl
    {
        // Dashboard tab controls
        private Label lblTitle;
        private Label lblWeekRange;
        private Label lblTotalHours;
        private Label lblHoursLabel;
        private Panel pnlChart;
        private Button btnPrevWeek;
        private Button btnNextWeek;
        private Button btnRefresh;
        private Panel pnlSummary;

        // Dashboard fonts
        private Font _fontTitle;
        private Font _fontWeekRange;
        private Font _fontTotalHours;
        private Font _fontHoursLabel;
        private Font _fontWeeklyTarget;
        private Font _fontLastWeekComparison;

        // Dashboard summary state
        private string _weeklyTargetText = "Target: 0.0% of 32.5 hours";
        private Color _weeklyTargetColor = Color.Gray;

        private string _lastWeekTitleText = "Last week: 0.0 hrs";
        private string _lastWeekDeltaText = "▲ 0.0 hours more";
        private Color _lastWeekDeltaColor = Color.FromArgb(0, 150, 0);

        // Tab control
        private TabControl tabControl;
        private TabPage tabDashboard;
        private TabPage tabSubmitted;  // ✅ NEW: Submitted tab
        private TabPage tabUnsubmitted;

        // Submitted tab controls - ✅ NEW
        private FlowLayoutPanel flowSubmitted;
        private Button btnRefreshSubmitted;
        private Font _fontSubmittedTitle;

        // Unsubmitted tab controls
        private FlowLayoutPanel flowUnsubmitted;
        private Button btnRefreshUnsubmitted;

        // Font fields
        private Font _fontUnsubmittedTitle;
        private Font _fontSectionHeader;
        private Font _fontSubject;
        private Font _fontTime;
        private Font _fontLabel;
        private Font _fontButton;
        private Font _fontButtonBold;

        private DateTime _currentWeekStart;
        private double[] _dailyHours;
        private string[] _dayLabels;
        private double _lastWeekTotal = 0;
        private bool _isInitialized = false;  // UI controls created
        private bool _isDisposed = false;

        // Cache for unsubmitted meetings
        private List<MeetingRecord> _cachedUnsubmittedMeetings = null;
        private DateTime _cacheExpiry = DateTime.MinValue;

        // PERF: Cache timezone lookup — FindSystemTimeZoneById scans the OS registry each call
        private static readonly TimeZoneInfo TorontoTz =
            TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");

        // Win32: hide scrollbar visually while keeping mouse-wheel scrolling functional
        [DllImport("user32.dll")] private static extern bool ShowScrollBar(IntPtr hWnd, int wBar, bool bShow);
        private const int SB_VERT = 1;

        private static void AttachAutoHideScrollbar(FlowLayoutPanel panel)
        {
            // Hide scrollbar once the handle exists
            panel.HandleCreated += (s, e) => ShowScrollBar(panel.Handle, SB_VERT, false);

            // Show on mouse-enter / any scroll activity, hide on mouse-leave
            panel.MouseEnter  += (s, e) => ShowScrollBar(panel.Handle, SB_VERT, true);
            panel.MouseLeave  += (s, e) => ShowScrollBar(panel.Handle, SB_VERT, false);
            panel.Scroll      += (s, e) => ShowScrollBar(panel.Handle, SB_VERT, true);

            // Mouse-wheel still scrolls even when bar is hidden
            panel.MouseWheel += (s, e) =>
            {
                ShowScrollBar(panel.Handle, SB_VERT, true);
                int newVal = Math.Max(panel.VerticalScroll.Minimum,
                             Math.Min(panel.VerticalScroll.Maximum,
                                      panel.VerticalScroll.Value - e.Delta));
                panel.VerticalScroll.Value = newVal;
            };
        }

        // Fixed pane width matches ThisAddIn.PaneFixedWidth.
        // TabControl chrome (borders + padding) eats ~8 px on each side.
        private const int PaneFixedWidth = 370;
        private const int TabChrome      = 8;

        /// <summary>
        /// Returns the usable card width inside a FlowLayoutPanel, even when the
        /// panel's parent tab page has never been selected (Width == 0).
        /// Falls back to a calculation based on the fixed pane width.
        /// </summary>
        private int GetFlowContentWidth(FlowLayoutPanel flow)
        {
            int clientW = flow.ClientSize.Width;
            if (clientW <= 0)
            {
                // Tab page not yet laid out — compute from the fixed pane width.
                // PaneFixedWidth − 2 × TabChrome − horizontal padding
                clientW = PaneFixedWidth - (TabChrome * 2) - flow.Padding.Horizontal;
            }
            return clientW - flow.Padding.Horizontal - SystemInformation.VerticalScrollBarWidth;
        }

        public ManageTimesheetPane()
        {
            this.AutoScaleMode = AutoScaleMode.None;
            InitializeComponent();
            _currentWeekStart = GetMondayOfCurrentWeek(DateTime.Now);
            _dailyHours = new double[7];
            _dayLabels = new[] { "Mo", "Tu", "We", "Th", "Fr", "Sa", "Su" };
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && !_isDisposed)
            {
                _isDisposed = true;
                _fontTitle?.Dispose();
                _fontWeekRange?.Dispose();
                _fontTotalHours?.Dispose();
                _fontHoursLabel?.Dispose();
                _fontWeeklyTarget?.Dispose();
                _fontLastWeekComparison?.Dispose();
                _fontSubmittedTitle?.Dispose();
                _fontUnsubmittedTitle?.Dispose();
                _fontSectionHeader?.Dispose();
                _fontSubject?.Dispose();
                _fontTime?.Dispose();
                _fontLabel?.Dispose();
                _fontButton?.Dispose();
                _fontButtonBold?.Dispose();
            }
            base.Dispose(disposing);
        }

        private void DisposeAndClearControls(Control.ControlCollection controls)
        {
            var list = new Control[controls.Count];
            controls.CopyTo(list, 0);
            controls.Clear();
            foreach (var c in list) c.Dispose();
        }

        protected override void OnVisibleChanged(EventArgs e)
        {
            base.OnVisibleChanged(e);
        }

        public async Task LoadDataAsync()
        {
            // Dashboard + Submitted are pure SQL — run them concurrently.
            // Both start synchronously on the UI thread (Invoke for "Loading…" labels),
            // then yield at cn.OpenAsync().ConfigureAwait(false), so the SQL work overlaps.
            var weeklyTask = LoadWeeklyDataAsync();
            var submittedTask = LoadSubmittedMeetingsAsync();

            try { await weeklyTask; }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"LoadWeeklyDataAsync failed: {ex.Message}"); }

            try { await submittedTask; }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"LoadSubmittedMeetingsAsync failed: {ex.Message}"); }

            // Unsubmitted accesses Outlook COM (STA) after its SQL — must run on
            // the UI thread, so it starts after the parallel SQL pair finishes.
            try { await LoadUnsubmittedMeetingsAsync(); }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"LoadUnsubmittedMeetingsAsync failed: {ex.Message}"); }
        }

        private void InitializeComponent()
        {
            if (_isInitialized) return;

            this.SuspendLayout();

            tabControl = new TabControl { Dock = DockStyle.Fill };

            tabDashboard = new TabPage("Dashboard");
            InitializeDashboardTab();

            tabSubmitted = new TabPage("Submitted Items");  // ✅ NEW
            InitializeSubmittedTab();

            tabUnsubmitted = new TabPage("Unsubmitted Items");
            InitializeUnsubmittedTab();

            tabControl.TabPages.Add(tabDashboard);
            tabControl.TabPages.Add(tabSubmitted);  // ✅ NEW
            tabControl.TabPages.Add(tabUnsubmitted);

            this.Controls.Add(tabControl);
            this.Name = "ManageTimesheetPane";
            this.Dock = DockStyle.Fill;

            _isInitialized = true;
            this.ResumeLayout(false);
        }

        private void InitializeDashboardTab()
        {
            _fontTitle = new Font("Segoe UI", 14, FontStyle.Bold);
            _fontWeekRange = new Font("Segoe UI", 10);
            _fontTotalHours = new Font("Segoe UI", 36, FontStyle.Bold);
            _fontHoursLabel = new Font("Segoe UI", 12);
            _fontWeeklyTarget = new Font("Segoe UI", 9);
            _fontLastWeekComparison = new Font("Segoe UI", 9);

            lblTitle = new Label
            {
                Text = "Weekly Hours",
                Font = _fontTitle,
                Location = new Point(15, 15),
                Size = new Size(320, 30),
                TextAlign = ContentAlignment.MiddleLeft
            };

            btnPrevWeek = new Button
            {
                Text = "◀",
                Location = new Point(15, 50),
                Size = new Size(30, 25)
            };
            btnPrevWeek.Click += BtnPrevWeek_Click;

            lblWeekRange = new Label
            {
                Text = "Loading...",
                Font = _fontWeekRange,
                Location = new Point(50, 50),
                Size = new Size(220, 25),
                TextAlign = ContentAlignment.MiddleCenter
            };

            btnNextWeek = new Button
            {
                Text = "▶",
                Location = new Point(275, 50),
                Size = new Size(30, 25)
            };
            btnNextWeek.Click += BtnNextWeek_Click;

            btnRefresh = new Button
            {
                Text = "↻",
                Location = new Point(310, 50),
                Size = new Size(30, 25)
            };
            btnRefresh.Click += BtnRefresh_Click;

            pnlChart = new Panel
            {
                Location = new Point(15, 90),
                Size = new Size(310, 180),
                BorderStyle = BorderStyle.FixedSingle
            };
            pnlChart.Paint += PnlChart_Paint;

            lblTotalHours = new Label
            {
                Text = "0.0",
                Font = _fontTotalHours,
                Location = new Point(15, 285),
                Size = new Size(180, 60),
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = Color.FromArgb(0, 120, 212)
            };

            lblHoursLabel = new Label
            {
                Text = "hours",
                Font = _fontHoursLabel,
                Location = new Point(200, 305),
                Size = new Size(125, 25),
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = Color.Gray
            };

            pnlSummary = new Panel
            {
                Location = new Point(15, 350),
                Size = new Size(310, 90),
                BackColor = Color.Transparent
            };
            pnlSummary.Paint += PnlSummary_Paint;

            tabDashboard.Controls.AddRange(new Control[]
            {
        lblTitle, btnPrevWeek, lblWeekRange, btnNextWeek, btnRefresh,
        pnlChart, lblTotalHours, lblHoursLabel,
        pnlSummary
            });
        }

        private async void BtnPrevWeek_Click(object sender, EventArgs e)
        {
            _currentWeekStart = _currentWeekStart.AddDays(-7);
            await LoadWeeklyDataAsync();
        }

        private async void BtnNextWeek_Click(object sender, EventArgs e)
        {
            _currentWeekStart = _currentWeekStart.AddDays(7);
            await LoadWeeklyDataAsync();
        }

        private async void BtnRefresh_Click(object sender, EventArgs e)
        {
            await LoadWeeklyDataAsync();
        }

        // ✅ NEW: Initialize Submitted tab
        private void InitializeSubmittedTab()
        {
            _fontSubmittedTitle = new Font("Segoe UI", 14, FontStyle.Bold);
            _fontSectionHeader  = new Font("Segoe UI", 11, FontStyle.Bold);
            _fontSubject        = new Font("Segoe UI", 9,  FontStyle.Bold);
            _fontTime           = new Font("Segoe UI", 8);
            _fontLabel          = new Font("Segoe UI", 10);
            _fontButton         = new Font("Segoe UI", 8);
            _fontButtonBold     = new Font("Segoe UI", 8,  FontStyle.Bold);

            // Top bar: title + refresh button at fixed height
            var topSubmitted = new Panel { Dock = DockStyle.Top, Height = 45 };

            var lblSubmittedTitle = new Label
            {
                Text = "Submitted",
                Font = _fontSubmittedTitle,
                Location = new Point(10, 8),
                Size = new Size(200, 30),
                TextAlign = ContentAlignment.MiddleLeft
            };

            btnRefreshSubmitted = new Button
            {
                Text = "↻",
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                Size = new Size(35, 25)
            };
            btnRefreshSubmitted.Location = new Point(topSubmitted.Width - 45, 10);
            btnRefreshSubmitted.Click += async (s, e) => await LoadSubmittedMeetingsAsync();
            topSubmitted.Controls.AddRange(new Control[] { lblSubmittedTitle, btnRefreshSubmitted });
            topSubmitted.Resize += (s, e) =>
                btnRefreshSubmitted.Location = new Point(topSubmitted.Width - 45, 10);

            // Flow panel fills the rest of the tab page
            flowSubmitted = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                Padding = new Padding(10, 5, 10, 0)
            };

            AttachAutoHideScrollbar(flowSubmitted);
            tabSubmitted.Controls.Add(flowSubmitted);   // Fill added first
            tabSubmitted.Controls.Add(topSubmitted);    // Top dock applied after
        }

        private void InitializeUnsubmittedTab()
        {
            _fontUnsubmittedTitle = new Font("Segoe UI", 14, FontStyle.Bold);

            var topUnsubmitted = new Panel { Dock = DockStyle.Top, Height = 45 };

            var lblUnsubmittedTitle = new Label
            {
                Text = "Unsubmitted",
                Font = _fontUnsubmittedTitle,
                Location = new Point(10, 8),
                Size = new Size(200, 30),
                TextAlign = ContentAlignment.MiddleLeft
            };

            btnRefreshUnsubmitted = new Button
            {
                Text = "↻",
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                Size = new Size(35, 25)
            };
            btnRefreshUnsubmitted.Location = new Point(topUnsubmitted.Width - 45, 10);
            btnRefreshUnsubmitted.Click += async (s, e) =>
            {
                _cachedUnsubmittedMeetings = null;
                _cacheExpiry = DateTime.MinValue;
                await LoadUnsubmittedMeetingsAsync();
            };
            topUnsubmitted.Controls.AddRange(new Control[] { lblUnsubmittedTitle, btnRefreshUnsubmitted });
            topUnsubmitted.Resize += (s, e) =>
                btnRefreshUnsubmitted.Location = new Point(topUnsubmitted.Width - 45, 10);

            flowUnsubmitted = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                Padding = new Padding(10, 5, 10, 0)
            };

            AttachAutoHideScrollbar(flowUnsubmitted);
            tabUnsubmitted.Controls.Add(flowUnsubmitted);
            tabUnsubmitted.Controls.Add(topUnsubmitted);
        }

        // ✅ Load submitted from database (shows: Today, Yesterday, Last Week + Cancel/Un-Ignore buttons)
        private async Task LoadSubmittedMeetingsAsync()
        {
            try
            {
                this.Invoke((MethodInvoker)delegate
                {
                    DisposeAndClearControls(flowSubmitted.Controls);
                    flowSubmitted.Controls.Add(new Label
                    {
                        Text = "Loading submitted events...",
                        Font = _fontLabel,
                        Width = GetFlowContentWidth(flowSubmitted),
                        Height = 30,
                        ForeColor = Color.Gray
                    });
                });

                var email = GetCurrentUserEmail();
                if (string.IsNullOrWhiteSpace(email)) throw new Exception("Unable to determine current user email.");

                var connString = ConfigurationManager.ConnectionStrings["OemsDatabase"]?.ConnectionString;
                if (string.IsNullOrWhiteSpace(connString)) throw new Exception("Database connection not configured.");

                var submittedMeetings = new List<MeetingRecord>();
                var ignoredMeetings = new List<MeetingRecord>();

                using (var cn = new SqlConnection(connString))
                {
                    await cn.OpenAsync().ConfigureAwait(false);
                    
                    // ✅ Load SUBMITTED records (status = 'submitted')
                    using (var cmd = new SqlCommand(@"
                        SELECT DISTINCT global_id, entry_id, subject, start_utc, end_utc,
                                job_code, activity_code, stage_code, hours_allocated
                        FROM db_owner.ytimesheet
                        WHERE user_name = @email AND status = 'submitted'
                          AND start_utc >= DATEADD(DAY, -30, CAST(GETDATE() AT TIME ZONE 'Eastern Standard Time' AS DATE))
                        ORDER BY start_utc DESC", cn))
                    {
                        cmd.Parameters.Add(new SqlParameter("@email", SqlDbType.NVarChar, 320) { Value = email });
                        using (var reader = await cmd.ExecuteReaderAsync().ConfigureAwait(false))
                        {
                            while (await reader.ReadAsync().ConfigureAwait(false))
                            {
                                var startTorontoTime = reader["start_utc"] is DateTime s ? s : DateTime.MinValue;
                                var endTorontoTime = reader["end_utc"] is DateTime e ? e : DateTime.MinValue;
                                
                                // ✅ CRITICAL FIX: Database stores Toronto time in start_utc column
                                // We need to convert it back to UTC for the MeetingRecord
                                var startUtc = TimeZoneInfo.ConvertTimeToUtc(startTorontoTime, TorontoTz);
                                var endUtc   = TimeZoneInfo.ConvertTimeToUtc(endTorontoTime,   TorontoTz);
                                
                                submittedMeetings.Add(new MeetingRecord
                                {
                                    GlobalId = reader["global_id"] as string ?? "",
                                    EntryId = reader["entry_id"] as string ?? "",
                                    Subject = reader["subject"] as string ?? "",
                                    StartUtc = startUtc,  // ✅ Convert back to UTC
                                    EndUtc = endUtc,      // ✅ Convert back to UTC
                                    ProgramCode = reader["job_code"] as string ?? "",
                                    ActivityCode = reader["activity_code"] as string ?? "",
                                    StageCode = reader["stage_code"] as string ?? "",
                                    UserDisplayName = email,
                                    Status = "submitted"
                                });
                            }
                        }
                    }

                    // ✅ Load IGNORED records (status = 'ignored')
                    using (var cmd = new SqlCommand(@"
                        SELECT DISTINCT global_id, entry_id, subject, start_utc, end_utc
                        FROM db_owner.ytimesheet
                        WHERE user_name = @email AND status = 'ignored'
                          AND start_utc >= DATEADD(DAY, -30, CAST(GETDATE() AT TIME ZONE 'Eastern Standard Time' AS DATE))
                        ORDER BY start_utc DESC", cn))
                    {
                        cmd.Parameters.Add(new SqlParameter("@email", SqlDbType.NVarChar, 320) { Value = email });
                        using (var reader = await cmd.ExecuteReaderAsync().ConfigureAwait(false))
                        {
                            while (await reader.ReadAsync().ConfigureAwait(false))
                            {
                                // DB stores Toronto local time in start_utc / end_utc (column is misnamed).
                                // Convert Toronto → UTC so StartTorontoTime round-trips back correctly,
                                // and the @start_utc param in CancelIgnoreTimesheetAsync matches the stored value.
                                var startTorontoTime = reader["start_utc"] is DateTime s ? s : DateTime.MinValue;
                                var endTorontoTime   = reader["end_utc"]   is DateTime e ? e : DateTime.MinValue;
                                var startUtc = TimeZoneInfo.ConvertTimeToUtc(startTorontoTime, TorontoTz);
                                var endUtc   = TimeZoneInfo.ConvertTimeToUtc(endTorontoTime,   TorontoTz);
                                ignoredMeetings.Add(new MeetingRecord
                                {
                                    GlobalId = reader["global_id"] as string ?? "",
                                    EntryId = reader["entry_id"] as string ?? "",
                                    Subject = reader["subject"] as string ?? "",
                                    StartUtc = startUtc,
                                    EndUtc   = endUtc,
                                    UserDisplayName = email,
                                    Status = "ignored"
                                });
                            }
                        }
                    }
                }

                // Merge submitted (grouped by GlobalId) and ignored into one unified list,
                // then group by date so both types appear together per date bucket.
                var groupedSubmitted = submittedMeetings
                    .GroupBy(m => m.GlobalId)
                    .Select(g => new SubmittedTabItem { IsIgnored = false, Records = g.ToList() })
                    .ToList();

                var groupedIgnored = ignoredMeetings
                    .Select(m => new SubmittedTabItem { IsIgnored = true, Records = new List<MeetingRecord> { m } })
                    .ToList();

                var allItems = groupedSubmitted.Concat(groupedIgnored).ToList();

                var today = DateTime.Today;

                var todayItems = allItems
                    .Where(i => i.StartTorontoTime.Date == today)
                    .OrderByDescending(i => i.StartTorontoTime).ToList();

                var yesterdayItems = allItems
                    .Where(i => i.StartTorontoTime.Date == today.AddDays(-1))
                    .OrderByDescending(i => i.StartTorontoTime).ToList();

                var lastWeekItems = allItems
                    .Where(i => i.StartTorontoTime.Date >= today.AddDays(-7) && i.StartTorontoTime.Date < today.AddDays(-1))
                    .OrderByDescending(i => i.StartTorontoTime).ToList();

                this.Invoke((MethodInvoker)delegate
                {
                    DisposeAndClearControls(flowSubmitted.Controls);

                    if (allItems.Count == 0)
                    {
                        flowSubmitted.Controls.Add(new Label
                        {
                            Text = "No submitted or ignored events found.",
                            Font = _fontLabel,
                            Size = new Size(GetFlowContentWidth(flowSubmitted), 40),
                            ForeColor = Color.Gray
                        });
                        return;
                    }

                    if (todayItems.Count > 0)     AddSubmittedTabSection("Today",     todayItems);
                    if (yesterdayItems.Count > 0) AddSubmittedTabSection("Yesterday", yesterdayItems);
                    if (lastWeekItems.Count > 0)  AddSubmittedTabSection("Last Week", lastWeekItems);

                    flowSubmitted.Controls.Add(new Label { Size = new Size(GetFlowContentWidth(flowSubmitted), 100), Text = "" });
                });
            }
            catch (Exception ex)
            {
                this.Invoke((MethodInvoker)delegate
                {
                    DisposeAndClearControls(flowSubmitted.Controls);
                    flowSubmitted.Controls.Add(new Label
                    {
                        Text = $"Error: {ex.Message}",
                        Font = _fontLabel,
                        Width = GetFlowContentWidth(flowSubmitted),
                        Height = 60,
                        ForeColor = Color.Red
                    });
                });
            }
        }

        // Unified wrapper: one entry per unique meeting in the submitted tab,
        // regardless of whether it is submitted (possibly multi-program) or ignored.
        private class SubmittedTabItem
        {
            public bool IsIgnored { get; set; }
            public List<MeetingRecord> Records { get; set; }
            public DateTime StartTorontoTime => Records[0].StartTorontoTime;
        }

        // Renders one date-bucket section (Today / Yesterday / Last Week) for the
        // submitted tab. Submitted and ignored cards are interleaved, sorted by time.
        private void AddSubmittedTabSection(string title, List<SubmittedTabItem> items)
        {
            int w = GetFlowContentWidth(flowSubmitted);

            // ── layout constants — identical to unsubmitted cards ─────────────
            const int subjectY  = 5;
            const int subjectH  = 24;
            const int timeY     = 30;
            const int timeH     = 20;
            const int btnY      = 53;   // immediately after time row
            const int btnH      = 27;
            const int btnGap    = 5;
            int panelHeight     = btnY + btnH + 6;

            flowSubmitted.Controls.Add(new Label
            {
                Text      = $"{title} ({items.Count})",
                Font      = _fontSectionHeader,
                Size      = new Size(w, 30),
                Margin    = new Padding(0, 5, 0, 0),
                ForeColor = Color.FromArgb(0, 120, 212),
                TextAlign = ContentAlignment.MiddleLeft
            });

            foreach (var item in items)
            {
                var first = item.Records[0];
                int bw = (w - 15 - btnGap) / 2;

                var p = new Panel
                {
                    Size        = new Size(w, panelHeight),
                    BorderStyle = BorderStyle.FixedSingle,
                    BackColor   = Color.FromArgb(250, 250, 250),
                    Margin      = new Padding(0, 0, 0, 6)
                };

                // Subject
                p.Controls.Add(new Label
                {
                    Text         = first.Subject,
                    Font         = _fontSubject,
                    Location     = new Point(5, subjectY),
                    Size         = new Size(w - 15, subjectH),
                    AutoEllipsis = true
                });

                // Date / duration
                p.Controls.Add(new Label
                {
                    Text      = $"{first.StartTorontoTime:MMM dd, HH:mm} ({(first.EndUtc - first.StartUtc).TotalHours:F1} hrs)",
                    Font      = _fontTime,
                    Location  = new Point(5, timeY),
                    Size      = new Size(w - 15, timeH),
                    ForeColor = Color.FromArgb(100, 100, 100)
                });

                // Buttons
                if (item.IsIgnored)
                {
                    // Ignored card: Un-Ignore only (full width)
                    var btnUnignore = new Button
                    {
                        Text      = "Un-Ignore",
                        Location  = new Point(5, btnY),
                        Size      = new Size(w - 15, btnH),
                        BackColor = Color.FromArgb(220, 220, 220),
                        ForeColor = Color.FromArgb(60, 60, 60),
                        FlatStyle = FlatStyle.Flat,
                        Font      = _fontButton
                    };
                    btnUnignore.FlatAppearance.BorderSize = 0;
                    var capturedItem = item;
                    btnUnignore.Click += async (s, e) => await CancelIgnoreSubmissionAsync(capturedItem.Records[0]);
                    p.Controls.Add(btnUnignore);
                }
                else
                {
                    // Submitted card: Cancel Submit (full width)
                    var btnCancel = new Button
                    {
                        Text      = "Cancel Submit",
                        Location  = new Point(5, btnY),
                        Size      = new Size(w - 15, btnH),
                        BackColor = Color.FromArgb(220, 100, 100),
                        ForeColor = Color.White,
                        FlatStyle = FlatStyle.Flat,
                        Font      = _fontButtonBold
                    };
                    btnCancel.FlatAppearance.BorderSize = 0;
                    var recordsToCancel = item.Records;
                    btnCancel.Click += async (s, e) => await CancelSubmissionAsync(recordsToCancel);
                    p.Controls.Add(btnCancel);
                }

                flowSubmitted.Controls.Add(p);
            }

            flowSubmitted.Controls.Add(new Label { Size = new Size(w, 10) });
        }

        // ✅ Cancel submission - delete timesheet
        private async Task CancelSubmissionAsync(List<MeetingRecord> meetings)
        {
            // ✅ NEW: Accept list of all records for this meeting
            if (meetings == null || meetings.Count == 0) 
            {
                System.Diagnostics.Debug.WriteLine("CancelSubmissionAsync: No meetings provided!");
                return;
            }

            var firstMeeting = meetings[0];
            
            System.Diagnostics.Debug.WriteLine($"CancelSubmissionAsync: Attempting to cancel {meetings.Count} record(s) for '{firstMeeting.Subject}'");
            
            if (MessageBox.Show(
                meetings.Count > 1
                    ? $"Cancel submission for {firstMeeting.Subject}?\n\nThis will delete {meetings.Count} program records.\nPrograms: {string.Join(", ", meetings.Select(m => m.ProgramCode))}\n\nThis action cannot be undone."
                    : $"Cancel submission for {firstMeeting.Subject}?\nProgram: {firstMeeting.ProgramCode}\nThis will remove the timesheet record.",
                "Cancel Submission", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) 
            {
                System.Diagnostics.Debug.WriteLine("CancelSubmissionAsync: User cancelled the dialog");
                return;
            }

            try
            {
                var email = GetCurrentUserEmail();
                
                // ✅ CRITICAL FIX: Delete ALL records for this meeting at once
                int deletedCount = 0;
                foreach (var record in meetings)
                {
                    System.Diagnostics.Debug.WriteLine($"CancelSubmissionAsync: Deleting record - Program:{record.ProgramCode}, GlobalId:{record.GlobalId}");
                    if (await DbWriter.DeleteTimesheetAsync(record)) 
                    {
                        deletedCount++;
                        System.Diagnostics.Debug.WriteLine($"CancelSubmissionAsync: Successfully deleted 1 record, total: {deletedCount}");
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"CancelSubmissionAsync: Failed to delete record for {record.ProgramCode}");
                    }
                }

                if (deletedCount > 0)
                {
                    System.Diagnostics.Debug.WriteLine($"CancelSubmissionAsync: Deletion complete, removing category and reloading");
                    RemoveTimesheetCategoryFromAppointment(firstMeeting.EntryId);
                    MessageBox.Show(
                        meetings.Count > 1
                            ? $"Deleted {deletedCount} program record(s) for this meeting."
                            : $"Deleted {deletedCount} timesheet record.",
                        "Cancelled");

                    // Refresh both tabs: item disappears from Submitted, reappears in Unsubmitted
                    _cachedUnsubmittedMeetings = null;
                    _cacheExpiry = DateTime.MinValue;
                    await LoadSubmittedMeetingsAsync();
                    await LoadUnsubmittedMeetingsAsync();
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("CancelSubmissionAsync: No records were deleted!");
                    MessageBox.Show("Failed to delete any records.", "Error");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"CancelSubmissionAsync: Exception occurred: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"CancelSubmissionAsync: Stack trace: {ex.StackTrace}");
                MessageBox.Show($"Error: {ex.Message}", "Error");
            }
        }

        // ✅ Cancel ignore - remove ignored status
        private async Task CancelIgnoreSubmissionAsync(MeetingRecord meeting)
        {
            System.Diagnostics.Debug.WriteLine($"CancelIgnoreSubmissionAsync: Attempting to un-ignore '{meeting.Subject}'");
            
            if (MessageBox.Show($"Remove ignore status from {meeting.Subject}?\nThis will allow it to appear in Unsubmitted list.",
                "Un-Ignore Meeting", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) 
            {
                System.Diagnostics.Debug.WriteLine("CancelIgnoreSubmissionAsync: User cancelled the dialog");
                return;
            }

            try
            {
                var email = GetCurrentUserEmail();
                var tempRec = new MeetingRecord
                {
                    GlobalId = meeting.GlobalId,
                    EntryId = meeting.EntryId,
                    StartUtc = meeting.StartUtc,
                    UserDisplayName = email
                };

                System.Diagnostics.Debug.WriteLine($"CancelIgnoreSubmissionAsync: Calling DbWriter.CancelIgnoreTimesheetAsync");
                if (await DbWriter.CancelIgnoreTimesheetAsync(tempRec))
                {
                    System.Diagnostics.Debug.WriteLine("CancelIgnoreSubmissionAsync: Successfully un-ignored, removing category and reloading");
                    RemoveTimesheetCategoryFromAppointment(meeting.EntryId);
                    MessageBox.Show("Ignore status removed!", "Un-Ignored");

                    // Invalidate unsubmitted cache so the item reappears
                    _cachedUnsubmittedMeetings = null;
                    _cacheExpiry = DateTime.MinValue;
                    await LoadSubmittedMeetingsAsync();
                    await LoadUnsubmittedMeetingsAsync();
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("CancelIgnoreSubmissionAsync: DbWriter.CancelIgnoreTimesheetAsync returned false");
                    MessageBox.Show("Failed to remove ignore status.", "Error");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"CancelIgnoreSubmissionAsync: Exception occurred: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"CancelIgnoreSubmissionAsync: Stack trace: {ex.StackTrace}");
                MessageBox.Show($"Error: {ex.Message}", "Error");
            }
        }

        private async Task SubmitMeetingAsync(MeetingRecord meeting)
        {
            Outlook.NameSpace ns = null;
            Outlook.AppointmentItem appt = null;
            Outlook.UserProperties ups = null;

            try
            {
                var email = GetCurrentUserEmail();
                if (string.IsNullOrWhiteSpace(email))
                {
                    MessageBox.Show("Unable to determine current user email.", "Error");
                    return;
                }

                var tempRec = new MeetingRecord
                {
                    GlobalId = meeting.GlobalId,
                    EntryId = meeting.EntryId,
                    StartUtc = meeting.StartUtc,
                    UserDisplayName = email
                };

                // ✅ Get ALL existing timesheets for this meeting (handles multi-program)
                var existingTimesheets = await DbWriter.GetAllTimesheetsForMeetingAsync(tempRec);

                if (existingTimesheets != null && existingTimesheets.Count > 0)
                {
                    if (existingTimesheets.Count > 1)
                    {
                        // Multi-program timesheet exists
                        var torontoTime = existingTimesheets[0].LastModifiedTorontoTime;
                        var programList = string.Join("\n", existingTimesheets.Select(t =>
                            $"  • {t.ProgramCode} ({(t.HoursAllocated ?? (t.EndUtc - t.StartUtc).TotalHours):F2} hrs)"));

                        var dialogResult = MessageBox.Show(
                            $"Multi-program timesheet already submitted for this meeting.\n\n" +
                            $"Previously submitted:\n{programList}\n" +
                            $"Submitted: {torontoTime:yyyy-MM-dd HH:mm} (Toronto Time)\n\n" +
                            $"Would you like to re-submit this timesheet?\n" +
                            $"(This will DELETE the existing entries and allow you to create new ones)",
                            "Multi-Program Timesheet Exists",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question);

                        if (dialogResult != DialogResult.Yes)
                            return;

                        foreach (var existing in existingTimesheets)
                            await DbWriter.DeleteTimesheetAsync(existing);
                    }
                    else
                    {
                        // Single-program timesheet exists
                        var existingTimesheet = existingTimesheets[0];
                        var torontoTime = existingTimesheet.LastModifiedTorontoTime;

                        var dialogResult = MessageBox.Show(
                            $"Timesheet already submitted for this meeting.\n\n" +
                            $"Previously submitted:\n" +
                            $"Program:  {existingTimesheet.ProgramCode}\n" +
                            $"Activity: {existingTimesheet.ActivityCode}\n" +
                            $"Stage:    {existingTimesheet.StageCode}\n" +
                            $"Submitted: {torontoTime:yyyy-MM-dd HH:mm} (Toronto Time)\n\n" +
                            $"Would you like to edit this timesheet?",
                            "Timesheet Already Submitted",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question);

                        if (dialogResult != DialogResult.Yes)
                            return;

                        await DbWriter.DeleteTimesheetAsync(existingTimesheet);
                    }
                }

                // Get initial values from first existing record or defaults
                var firstExisting = existingTimesheets?.FirstOrDefault();
                string currProgram = firstExisting?.ProgramCode ?? "";
                string currActivity = firstExisting?.ActivityCode ?? "";
                string currStage = firstExisting?.StageCode ?? "";

                var duration = (meeting.EndUtc - meeting.StartUtc).TotalHours;
                var source = firstExisting != null ? "ManageTimesheet_Update" : "ManageTimesheet_Submit";

                using (var dlg = new ProgramPickerForm(currProgram, currActivity, currStage, duration))
                {
                    if (dlg.ShowDialog() != DialogResult.OK)
                        return;

                    try
                    {
                        ns = Globals.ThisAddIn.Application.Session;
                        appt = ns.GetItemFromID(meeting.EntryId) as Outlook.AppointmentItem;

                        // Collect recipients
                        string recipients = "";
                        if (appt != null)
                        {
                            try { recipients = GetAllRecipients(appt); }
                            catch (Exception recipEx)
                            {
                                System.Diagnostics.Debug.WriteLine($"Failed to get recipients: {recipEx.Message}");
                            }
                        }

                        if (dlg.IsMultiProgram && dlg.ProgramAllocations.Count >= 1)
                        {
                            // ✅ Multi-program mode
                            var allocations = new List<ProgramAllocation>();
                            double additionalHours = dlg.ProgramAllocations.Sum(p => p.Hours);
                            double originalProgramHours = duration - additionalHours;

                            // Original program first
                            allocations.Add(new ProgramAllocation
                            {
                                ProgramCode = dlg.ProgramCode,
                                ActivityCode = dlg.ActivityCode,
                                StageCode = dlg.StageCode,
                                Hours = originalProgramHours
                            });
                            allocations.AddRange(dlg.ProgramAllocations);

                            foreach (var allocation in allocations)
                            {
                                var rec = new MeetingRecord
                                {
                                    Source = source,
                                    EntryId = meeting.EntryId,
                                    GlobalId = meeting.GlobalId,
                                    Subject = meeting.Subject,
                                    StartUtc = meeting.StartUtc,
                                    EndUtc = meeting.EndUtc,
                                    HoursAllocated = allocation.Hours,
                                    ProgramCode = allocation.ProgramCode,
                                    ActivityCode = allocation.ActivityCode,
                                    StageCode = allocation.StageCode,
                                    UserDisplayName = email,
                                    LastModifiedUtc = DateTime.UtcNow,
                                    IsRecurring = meeting.IsRecurring,
                                    Recipients = recipients,
                                    Status = "submitted"
                                };
                                await DbWriter.UpsertAsync(rec);
                            }

                            if (appt != null)
                                ApplyCategoryToAppointment(appt, "Timesheet Submitted", Outlook.OlCategoryColor.olCategoryColorPeach);

                            MessageBox.Show(
                                $"Timesheet submitted for {allocations.Count} programs!\n\n" +
                                $"Original Program:\n  • {allocations[0].ProgramCode}: {allocations[0].Hours:F2} hrs\n\n" +
                                $"Additional Programs:\n" +
                                string.Join("\n", allocations.Skip(1).Select(a => $"  • {a.ProgramCode}: {a.Hours:F2} hrs")),
                                "Multi-Program Timesheet Submitted",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
                        }
                        else
                        {
                            // ✅ Single-program mode
                            string programCode, activityCode, stageCode;

                            if (dlg.IsMultiProgram && dlg.ProgramAllocations.Count == 1)
                            {
                                var singleAllocation = dlg.ProgramAllocations[0];
                                programCode = singleAllocation.ProgramCode ?? "";
                                activityCode = singleAllocation.ActivityCode ?? "";
                                stageCode = singleAllocation.StageCode ?? "";
                            }
                            else
                            {
                                programCode = dlg.ProgramCode ?? "";
                                activityCode = dlg.ActivityCode ?? "";
                                stageCode = dlg.StageCode ?? "";
                            }

                            if (appt != null)
                            {
                                ups = appt.UserProperties;
                                AddOrSetTextProp(ups, "ProgramCode", programCode);
                                AddOrSetTextProp(ups, "ActivityCode", activityCode);
                                AddOrSetTextProp(ups, "StageCode", stageCode);
                            }

                            var rec = new MeetingRecord
                            {
                                Source = source,
                                EntryId = meeting.EntryId,
                                GlobalId = meeting.GlobalId,
                                Subject = meeting.Subject,
                                StartUtc = meeting.StartUtc,
                                EndUtc = meeting.EndUtc,
                                ProgramCode = programCode,
                                ActivityCode = activityCode,
                                StageCode = stageCode,
                                UserDisplayName = email,
                                LastModifiedUtc = DateTime.UtcNow,
                                IsRecurring = meeting.IsRecurring,
                                Recipients = recipients,
                                Status = "submitted"
                            };

                            await DbWriter.UpsertAsync(rec);

                            if (appt != null)
                                ApplyCategoryToAppointment(appt, "Timesheet Submitted", Outlook.OlCategoryColor.olCategoryColorPeach);

                            var actionText = firstExisting != null ? "updated" : "submitted";
                            MessageBox.Show(
                                $"Timesheet {actionText} successfully!\n\n" +
                                $"Program:  {programCode}\n" +
                                $"Activity: {activityCode}\n" +
                                $"Stage:    {stageCode}",
                                $"Timesheet {(firstExisting != null ? "Updated" : "Submitted")}");
                        }

                        // Refresh the list
                        _cachedUnsubmittedMeetings = null;
                        _cacheExpiry = DateTime.MinValue;
                        await LoadUnsubmittedMeetingsAsync();
                    }
                    finally
                    {
                        if (ups != null) { Marshal.ReleaseComObject(ups); ups = null; }
                        if (appt != null) { Marshal.ReleaseComObject(appt); appt = null; }
                        if (ns != null) { Marshal.ReleaseComObject(ns); ns = null; }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SubmitMeetingAsync failed: {ex.Message}");
                MessageBox.Show($"Failed to submit timesheet: {ex.Message}", "Error");
            }
        }

        private async void IgnoreMeeting(MeetingRecord meeting, Panel panel)
        {
            try
            {
                var result = MessageBox.Show(
                    $"Permanently ignore this meeting?\n\n" +
                    $"Subject: {meeting.Subject}\n" +
                    $"Date: {meeting.StartTorontoTime:MMM dd, yyyy HH:mm}\n\n" +
                    $"This will permanently hide it from the unsubmitted list.",
                    "Ignore Meeting",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    var email = GetCurrentUserEmail();
                    meeting.UserDisplayName = email;

                    var success = await DbWriter.IgnoreTimesheetAsync(meeting);

                    System.Diagnostics.Debug.WriteLine($"IgnoreMeeting: DbWriter.IgnoreTimesheetAsync returned: {success}");

                    if (success)
                    {
                        System.Diagnostics.Debug.WriteLine($"🔍 IgnoreMeeting: meeting.IsRecurringOccurrence = {meeting.IsRecurringOccurrence}");

                        if (!meeting.IsRecurringOccurrence)
                        {
                            Microsoft.Office.Interop.Outlook.NameSpace ns = null;
                            Microsoft.Office.Interop.Outlook.AppointmentItem appt = null;
                            try
                            {
                                ns = Globals.ThisAddIn.Application.Session;
                                appt = ns.GetItemFromID(meeting.EntryId) as Microsoft.Office.Interop.Outlook.AppointmentItem;
                                if (appt != null)
                                {
                                    // Only apply category if not a recurring series
                                    if (appt.RecurrenceState != Microsoft.Office.Interop.Outlook.OlRecurrenceState.olApptMaster &&
                                        appt.RecurrenceState != Microsoft.Office.Interop.Outlook.OlRecurrenceState.olApptOccurrence)
                                    {
                                        ApplyCategoryToAppointment(appt, "Timesheet Ignored", Microsoft.Office.Interop.Outlook.OlCategoryColor.olCategoryColorDarkGray);
                                        System.Diagnostics.Debug.WriteLine($"✅ Applied 'Timesheet Ignored' category to non-recurring meeting: {appt.Subject}");
                                    }
                                }
                            }
                            finally
                            {
                                if (appt != null)
                                {
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(appt);
                                    appt = null;
                                }
                                if (ns != null)
                                {
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ns);
                                    ns = null;
                                }
                            }
                        }

                        System.Diagnostics.Debug.WriteLine($"Permanently ignored meeting: {meeting.Subject}");

                        // Clear cache and reload the unsubmitted meetings list
                        _cachedUnsubmittedMeetings = null;
                        _cacheExpiry = DateTime.MinValue;
                        await LoadUnsubmittedMeetingsAsync();
                    }
                    else
                    {
                        MessageBox.Show("Failed to ignore meeting. Please try again.", "Error");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"IgnoreMeeting failed: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                MessageBox.Show($"Failed to ignore meeting: {ex.Message}", "Error");
            }
        }

        // ✅ Helper to apply category to appointment with custom color
        private void ApplyCategoryToAppointment(Microsoft.Office.Interop.Outlook.AppointmentItem appt, string categoryName, Microsoft.Office.Interop.Outlook.OlCategoryColor categoryColor)
        {
            Microsoft.Office.Interop.Outlook.NameSpace session = null;
            Microsoft.Office.Interop.Outlook.Categories categories = null;
            Microsoft.Office.Interop.Outlook.Category existingCategory = null;

            try
            {
                if (appt.RecurrenceState == Microsoft.Office.Interop.Outlook.OlRecurrenceState.olApptOccurrence)
                {
                    System.Diagnostics.Debug.WriteLine($"⚠️ SKIPPING category for recurring meeting occurrence: {appt.Subject}");
                    return;
                }

                session = Globals.ThisAddIn.Application.Session;
                categories = session.Categories;
                foreach (Microsoft.Office.Interop.Outlook.Category c in categories)
                {
                    if (c.Name == categoryName) { existingCategory = c; break; }
                }

                if (existingCategory == null)
                {
                    categories.Add(categoryName, categoryColor);
                }
                else if (existingCategory.Color != categoryColor)
                {
                    categories.Remove(categoryName);
                    categories.Add(categoryName, categoryColor);
                }

                // Remove any existing timesheet categories first (submitted or ignored)
                if (!string.IsNullOrEmpty(appt.Categories))
                {
                    var existingCategories = appt.Categories.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(c => c.Trim())
                        .Where(c => !c.Equals("Timesheet Submitted", StringComparison.OrdinalIgnoreCase) &&
                                   !c.Equals("Timesheet Ignored", StringComparison.OrdinalIgnoreCase))
                        .ToList();

                    existingCategories.Add(categoryName);
                    appt.Categories = string.Join(", ", existingCategories);
                }
                else
                {
                    appt.Categories = categoryName;
                }

                appt.Save();
                System.Diagnostics.Debug.WriteLine($"ApplyCategoryToAppointment: Set categories to '{appt.Categories}'");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ApplyCategoryToAppointment failed: {ex.Message}");
            }
            finally
            {
                if (existingCategory != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(existingCategory);
                    existingCategory = null;
                }
                if (categories != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(categories);
                    categories = null;
                }
                if (session != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(session);
                    session = null;
                }
            }
        }

        /// <summary>
        /// Finds the Outlook appointment by EntryId and strips all timesheet-related
        /// categories ("Timesheet Submitted", "Timesheet Ignored").
        /// Safe to call even when the appointment no longer exists in the calendar.
        /// </summary>
        private void RemoveTimesheetCategoryFromAppointment(string entryId)
        {
            Outlook.NameSpace ns = null;
            Outlook.AppointmentItem appt = null;
            try
            {
                ns = Globals.ThisAddIn.Application.Session;
                appt = ns.GetItemFromID(entryId) as Outlook.AppointmentItem;

                if (appt == null)
                {
                    System.Diagnostics.Debug.WriteLine("RemoveTimesheetCategory: Appointment not found for EntryId");
                    return;
                }

                if (appt.RecurrenceState == Outlook.OlRecurrenceState.olApptOccurrence)
                {
                    System.Diagnostics.Debug.WriteLine($"RemoveTimesheetCategory: Skipping recurring occurrence: {appt.Subject}");
                    return;
                }

                if (string.IsNullOrEmpty(appt.Categories))
                    return;

                var remaining = appt.Categories
                    .Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(c => c.Trim())
                    .Where(c => !c.Equals("Timesheet Submitted", StringComparison.OrdinalIgnoreCase) &&
                                !c.Equals("Timesheet Ignored", StringComparison.OrdinalIgnoreCase) &&
                                !c.Equals("Multi-Program Timesheet Submitted", StringComparison.OrdinalIgnoreCase))
                    .ToList();

                appt.Categories = remaining.Count > 0 ? string.Join(", ", remaining) : null;
                appt.Save();
                System.Diagnostics.Debug.WriteLine($"RemoveTimesheetCategory: Categories now '{appt.Categories}'");
            }
            catch (Exception ex)
            {
                // Appointment may have been deleted from calendar — not critical
                System.Diagnostics.Debug.WriteLine($"RemoveTimesheetCategory failed: {ex.Message}");
            }
            finally
            {
                if (appt != null) Marshal.ReleaseComObject(appt);
                if (ns != null) Marshal.ReleaseComObject(ns);
            }
        }

        // ✅ Helper method to add or set user property
        private static void AddOrSetTextProp(Microsoft.Office.Interop.Outlook.UserProperties ups, string name, string value)
        {
            Microsoft.Office.Interop.Outlook.UserProperty up = null;
            try
            {
                up = ups.Find(name);
                if (up == null)
                {
                    up = ups.Add(name, Microsoft.Office.Interop.Outlook.OlUserPropertyType.olText, false, Type.Missing);
                }
                up.Value = value ?? string.Empty;
            }
            finally
            {
                if (up != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(up);
                    up = null;
                }
            }
        }
        private DateTime GetMondayOfCurrentWeek(DateTime date)
        {
            int diff = (7 + (date.DayOfWeek - DayOfWeek.Monday)) % 7;
            return date.AddDays(-1 * diff).Date;
        }

        private void PnlSummary_Paint(object sender, PaintEventArgs e)
        {
            var g = e.Graphics;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

            int x = 2;
            int y = 2;
            int w = pnlSummary.ClientSize.Width - 4;

            // Use NoClip so text is never sliced mid-character even at high DPI
            using (var sf = new StringFormat { FormatFlags = StringFormatFlags.NoClip })
            using (var targetBrush = new SolidBrush(_weeklyTargetColor))
            using (var grayBrush   = new SolidBrush(Color.Gray))
            using (var deltaBrush  = new SolidBrush(_lastWeekDeltaColor))
            {
                int h1 = TextRenderer.MeasureText(_weeklyTargetText,  _fontWeeklyTarget).Height;
                int h2 = TextRenderer.MeasureText(_lastWeekTitleText, _fontLastWeekComparison).Height;
                int h3 = TextRenderer.MeasureText(_lastWeekDeltaText, _fontLastWeekComparison).Height;

                g.DrawString(_weeklyTargetText,  _fontWeeklyTarget,      targetBrush, new RectangleF(x, y,      w, h1 + 4), sf);
                y += h1 + 2;
                g.DrawString(_lastWeekTitleText, _fontLastWeekComparison, grayBrush,  new RectangleF(x, y,      w, h2 + 4), sf);
                y += h2 + 2;
                g.DrawString(_lastWeekDeltaText, _fontLastWeekComparison, deltaBrush, new RectangleF(x, y,      w, h3 + 4), sf);
            }
        }
        private async Task LoadWeeklyDataAsync()
        {
            try
            {
                var weekEnd = _currentWeekStart.AddDays(6);

                this.Invoke((MethodInvoker)delegate
                {
                    lblWeekRange.Text = $"{_currentWeekStart:MMM dd} - {weekEnd:MMM dd, yyyy}";
                });

                var email = GetCurrentUserEmail();
                if (string.IsNullOrWhiteSpace(email))
                {
                    this.Invoke((MethodInvoker)delegate { MessageBox.Show("Unable to determine current user email.", "Error"); });
                    return;
                }

                var connString = ConfigurationManager.ConnectionStrings["OemsDatabase"]?.ConnectionString;
                if (string.IsNullOrWhiteSpace(connString))
                {
                    this.Invoke((MethodInvoker)delegate { MessageBox.Show("Database connection not configured.", "Error"); });
                    return;
                }

                Array.Clear(_dailyHours, 0, _dailyHours.Length);
                _lastWeekTotal = 0;

                var lastWeekStart = _currentWeekStart.AddDays(-7);

                using (var cn = new SqlConnection(connString))
                {
                    await cn.OpenAsync().ConfigureAwait(false);

                    // Current week
                    using (var cmd = new SqlCommand("dbo.Timesheet_GetWeeklyData", cn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@email", email);
                        cmd.Parameters.AddWithValue("@weekStart", _currentWeekStart);
                        cmd.Parameters.AddWithValue("@weekEnd", _currentWeekStart.AddDays(6));

                        using (var reader = await cmd.ExecuteReaderAsync().ConfigureAwait(false))
                        {
                            while (await reader.ReadAsync().ConfigureAwait(false))
                            {
                                var workDate = reader.GetDateTime(0);
                                double hours = 0;
                                if (!reader.IsDBNull(1))
                                {
                                    hours = Convert.ToDouble(reader[1]);
                                }
                                int dayIndex = (int)((workDate - _currentWeekStart).TotalDays);
                                if (dayIndex >= 0 && dayIndex < 7)
                                    _dailyHours[dayIndex] = hours;
                            }
                        }
                    }

                    // Last week — reuse the same open connection (no second SqlConnection needed)
                    using (var cmd = new SqlCommand("dbo.Timesheet_GetWeeklyData", cn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@email", email);
                        cmd.Parameters.AddWithValue("@weekStart", lastWeekStart);
                        cmd.Parameters.AddWithValue("@weekEnd", lastWeekStart.AddDays(6));

                        using (var reader = await cmd.ExecuteReaderAsync().ConfigureAwait(false))
                        {
                            while (await reader.ReadAsync().ConfigureAwait(false))
                            {
                                if (!reader.IsDBNull(1))
                                {
                                    _lastWeekTotal += Convert.ToDouble(reader[1]);
                                }
                            }
                        }
                    }
                }

                var totalHours = _dailyHours.Sum();
                double targetHours = 32.5;
                double targetPercentage = (totalHours / targetHours) * 100.0;
                double weekDifference = totalHours - _lastWeekTotal;

                string weekComparison = weekDifference >= 0
                    ? $"▲ {Math.Abs(weekDifference):F1} hours more than last week"
                    : $"▼ {Math.Abs(weekDifference):F1} hours less than last week";

                Color comparisonColor = weekDifference >= 0
                    ? Color.FromArgb(0, 150, 0)
                    : Color.FromArgb(200, 0, 0);

                // ✅ CRITICAL FIX: Marshal ALL UI updates to UI thread
                this.Invoke((MethodInvoker)delegate
                {
                    lblTotalHours.Text = totalHours.ToString("F1");

                    _weeklyTargetText = $"Target: {targetPercentage:F1}% of {targetHours} hours";
                    _weeklyTargetColor = targetPercentage >= 100
                        ? Color.FromArgb(0, 150, 0)
                        : Color.FromArgb(100, 100, 100);

                    _lastWeekTitleText = $"Last week: {_lastWeekTotal:F1} hrs";
                    _lastWeekDeltaText = weekComparison;
                    _lastWeekDeltaColor = comparisonColor;

                    pnlSummary.Invalidate();
                    pnlChart.Invalidate();
                });
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"LoadWeeklyDataAsync failed: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");

                // ✅ FIX: Show error on UI thread
                this.Invoke((MethodInvoker)delegate
                {
                    MessageBox.Show($"Failed to load timesheet data: {ex.Message}", "Error");
                });
            }
        }

        private void ShowLoadingMessage(string message)
        {
            this.Invoke((MethodInvoker)delegate
            {
                DisposeAndClearControls(flowUnsubmitted.Controls);
                flowUnsubmitted.Controls.Add(new Label
                {
                    Text = message,
                    Font = _fontLabel,
                    Width = GetFlowContentWidth(flowUnsubmitted),
                    Height = 30,
                    ForeColor = Color.Gray
                });
            });
        }

        private void ShowErrorMessage(string message)
        {
            this.Invoke((MethodInvoker)delegate
            {
                DisposeAndClearControls(flowUnsubmitted.Controls);
                flowUnsubmitted.Controls.Add(new Label
                {
                    Text = message,
                    Font = _fontLabel,
                    Width = GetFlowContentWidth(flowUnsubmitted),
                    Height = 60,
                    ForeColor = Color.Red
                });
            });
        }

        private void PnlChart_Paint(object sender, PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            g.Clear(Color.White);

            const int leftMargin = 30;
            const int chartWidth = 185;
            const int chartHeight = 130;
            const int chartX = leftMargin;
            const int chartY = 15;
            const int barWidth = 20;
            const int spacing = 5;
            const double maxHours = 8.0;

            using (var font = new Font("Segoe UI", 8))
            using (var pen = new Pen(Color.LightGray, 1) { DashStyle = System.Drawing.Drawing2D.DashStyle.Dot })
            {
                for (int h = 0; h <= 8; h += 2)
                {
                    int y = chartY + chartHeight - (int)((h / maxHours) * chartHeight);
                    g.DrawString(h.ToString(), font, Brushes.Gray, new PointF(8, y - 7));

                    if (h > 0)
                        g.DrawLine(pen, chartX, y, chartX + chartWidth, y);
                }
            }

            for (int i = 0; i < 7; i++)
            {
                int x = chartX + (i * (barWidth + spacing));

                using (var grayBrush = new SolidBrush(Color.FromArgb(220, 220, 220)))
                {
                    g.FillRectangle(grayBrush, x, chartY, barWidth, chartHeight);
                }

                double hours = _dailyHours[i];
                if (hours > 0)
                {
                    int filledHeight = Math.Min((int)((hours / maxHours) * chartHeight), chartHeight);
                    int filledY = chartY + chartHeight - filledHeight;

                    using (var blueBrush = new SolidBrush(Color.FromArgb(0, 120, 212)))
                    {
                        g.FillRectangle(blueBrush, x, filledY, barWidth, filledHeight);
                    }
                }

                using (var font = new Font("Segoe UI", 8, FontStyle.Bold))
                {
                    int xOffset = 2;
                    var dayText = _dayLabels[i];
                    var textSize = g.MeasureString(dayText, font);
                    g.DrawString(dayText, font, Brushes.Black,
                        new PointF(x + (barWidth - textSize.Width) / 2, chartY + chartHeight + 8));
                }
            }
        }

        private string GetCurrentUserEmail()
        {
            // Use cached value from ThisAddIn if available (avoids COM calls entirely)
            var cached = Globals.ThisAddIn?.GetCachedUserEmail();
            if (!string.IsNullOrWhiteSpace(cached))
                return cached;

            // Fallback: full COM lookup (first call only)
            Microsoft.Office.Interop.Outlook.NameSpace session = null;
            Microsoft.Office.Interop.Outlook.Recipient currentUser = null;
            Microsoft.Office.Interop.Outlook.AddressEntry addrEntry = null;
            Microsoft.Office.Interop.Outlook.PropertyAccessor pa = null;

            try
            {
                session = Globals.ThisAddIn?.Application?.Session;
                currentUser = session?.CurrentUser;
                addrEntry = currentUser?.AddressEntry;

                if (addrEntry != null)
                {
                    if ("EX".Equals(addrEntry.Type, StringComparison.OrdinalIgnoreCase))
                    {
                        const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
                        pa = addrEntry.PropertyAccessor;
                        var smtp = pa.GetProperty(PR_SMTP_ADDRESS) as string;
                        if (!string.IsNullOrWhiteSpace(smtp)) return smtp;
                    }
                    if (!string.IsNullOrWhiteSpace(addrEntry.Address)) return addrEntry.Address;
                }
                return currentUser?.Name ?? string.Empty;
            }
            catch { return string.Empty; }
            finally
            {
                if (pa != null) { System.Runtime.InteropServices.Marshal.ReleaseComObject(pa); pa = null; }
                if (addrEntry != null) { System.Runtime.InteropServices.Marshal.ReleaseComObject(addrEntry); addrEntry = null; }
                if (currentUser != null) { System.Runtime.InteropServices.Marshal.ReleaseComObject(currentUser); currentUser = null; }
                if (session != null) { System.Runtime.InteropServices.Marshal.ReleaseComObject(session); session = null; }
            }
        }

        private async Task<List<MeetingRecord>> GetUnsubmittedMeetingsFromOutlookAsync(string email)
        {
            var unsubmittedMeetings = new List<MeetingRecord>();
            var connString = ConfigurationManager.ConnectionStrings["OemsDatabase"]?.ConnectionString;
            if (string.IsNullOrWhiteSpace(connString))
                throw new InvalidOperationException("Database connection not configured.");

            var nowTorontoTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, TorontoTz);
            // Use midnight (Date) so the window matches the Outlook filter exactly
            var startDateTorontoTime = nowTorontoTime.Date.AddDays(-7);

            // Fetch BOTH global_id and entry_id so we can match against either ID
            // (Outlook may return GlobalAppointmentID or fall back to EntryID)
            var submittedOrIgnoredKeys = new HashSet<string>();
            using (var cn = new SqlConnection(connString))
            {
                // NOTE: No ConfigureAwait(false) here — this method accesses Outlook COM objects
                // after the SQL section. COM Outlook objects are STA-bound; resuming on an MTA
                // thread-pool thread forces inter-apartment marshalling on every property access
                // (~1-2 ms each × hundreds of calendar items = visible delay).
                await cn.OpenAsync();
                using (var cmd = new SqlCommand(@"
                    SELECT DISTINCT global_id, entry_id, start_utc
                    FROM db_owner.ytimesheet
                    WHERE user_name = @email
                      AND status IN ('submitted', 'ignored')
                      AND start_utc >= @startDate", cn))
                {
                    cmd.Parameters.Add(new SqlParameter("@email", SqlDbType.NVarChar, 320) { Value = email });
                    cmd.Parameters.Add(new SqlParameter("@startDate", SqlDbType.DateTime2) { Value = startDateTorontoTime });

                    using (var reader = await cmd.ExecuteReaderAsync())
                    {
                        while (await reader.ReadAsync())
                        {
                            var dbGlobalId = reader["global_id"] as string ?? "";
                            var dbEntryId  = reader["entry_id"]  as string ?? "";
                            var startTime  = reader.GetDateTime(reader.GetOrdinal("start_utc"));
                            var dateStr    = $"{startTime:yyyy-MM-dd}";

                            // Add keys for BOTH IDs so either will match the Outlook check
                            if (!string.IsNullOrWhiteSpace(dbGlobalId))
                                submittedOrIgnoredKeys.Add($"{dbGlobalId}|{dateStr}");
                            if (!string.IsNullOrWhiteSpace(dbEntryId))
                                submittedOrIgnoredKeys.Add($"{dbEntryId}|{dateStr}");
                        }
                    }
                }
            }

            // ── Outlook COM loop ──────────────────────────────────────────
            // Runs directly on the UI (STA) thread — NO Task.Run.
            // Outlook COM objects are STA-bound; calling from an MTA thread
            // forces cross-apartment marshaling (~1-2 ms per property access
            // × 7 properties × 200 items ≈ 1.4 s of pure overhead).
            // Running on STA avoids all marshaling; the loop takes ~100-200 ms.
            Microsoft.Office.Interop.Outlook.NameSpace ns = null;
            Microsoft.Office.Interop.Outlook.MAPIFolder calFolder = null;
            Microsoft.Office.Interop.Outlook.Items items = null;
            Microsoft.Office.Interop.Outlook.Items filteredItems = null;

            try
            {
                ns = Globals.ThisAddIn.Application.Session;
                calFolder = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
                items = calFolder.Items;
                items.Sort("[Start]", false);
                items.IncludeRecurrences = true;

                var startDate = DateTime.Today.AddDays(-7);
                var endDate = DateTime.Today.AddDays(1);
                var filter = $"[Start] >= '{startDate:g}' AND [Start] <= '{endDate:g}'";
                filteredItems = items.Restrict(filter);

                // Use GetFirst/GetNext — avoids .Count which forces Outlook to
                // enumerate ALL expanded recurrences upfront (very expensive).
                object rawItem = filteredItems.GetFirst();
                int count = 0;
                while (rawItem != null && count < 200)
                {
                    var appt = rawItem as Outlook.AppointmentItem;
                    if (appt != null)
                    {
                        try
                        {
                            var recurrenceState = appt.RecurrenceState;
                            if (recurrenceState != Outlook.OlRecurrenceState.olApptMaster)
                            {
                                string globalAppointmentId = "";
                                try { globalAppointmentId = appt.GlobalAppointmentID ?? ""; } catch { }
                                string entryId = appt.EntryID ?? "";
                                string globalId = !string.IsNullOrWhiteSpace(globalAppointmentId)
                                    ? globalAppointmentId : entryId;

                                if (!string.IsNullOrWhiteSpace(globalId))
                                {
                                    var apptStartUtc = appt.StartUTC;
                                    var dateStr = TimeZoneInfo.ConvertTimeFromUtc(apptStartUtc, TorontoTz).ToString("yyyy-MM-dd");

                                    bool alreadyProcessed =
                                        (!string.IsNullOrWhiteSpace(globalAppointmentId) &&
                                         submittedOrIgnoredKeys.Contains($"{globalAppointmentId}|{dateStr}")) ||
                                        (!string.IsNullOrWhiteSpace(entryId) &&
                                         submittedOrIgnoredKeys.Contains($"{entryId}|{dateStr}"));

                                    if (!alreadyProcessed &&
                                        appt.MeetingStatus != Outlook.OlMeetingStatus.olMeetingCanceled &&
                                        appt.MeetingStatus != Outlook.OlMeetingStatus.olMeetingReceivedAndCanceled)
                                    {
                                        unsubmittedMeetings.Add(new MeetingRecord
                                        {
                                            EntryId = entryId,
                                            GlobalId = globalId,
                                            Subject = appt.Subject ?? "",
                                            StartUtc = apptStartUtc,
                                            EndUtc = appt.EndUTC,
                                            UserDisplayName = email,
                                            IsRecurringOccurrence = recurrenceState == Outlook.OlRecurrenceState.olApptOccurrence
                                        });
                                    }
                                }
                            }
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(appt);
                        }
                    }
                    else if (rawItem != null)
                    {
                        Marshal.ReleaseComObject(rawItem);
                    }

                    rawItem = filteredItems.GetNext();
                    count++;
                }
            }
            finally
            {
                if (filteredItems != null) Marshal.ReleaseComObject(filteredItems);
                if (items != null) Marshal.ReleaseComObject(items);
                if (calFolder != null) Marshal.ReleaseComObject(calFolder);
                if (ns != null) Marshal.ReleaseComObject(ns);
            }

            return unsubmittedMeetings;
        }

        // ✅ Load unsubmitted from Outlook (shows: Today, Yesterday, Last Week + Submit/Ignore buttons)
        private async Task LoadUnsubmittedMeetingsAsync()
        {
            try
            {
                ShowLoadingMessage("Loading unsubmitted events...");

                var email = GetCurrentUserEmail();
                if (string.IsNullOrWhiteSpace(email))
                {
                    ShowErrorMessage("Unable to determine current user email.");
                    return;
                }

                // ✅ Use cached data if fresh (< 2 minutes old)
                List<MeetingRecord> unsubmittedMeetings;
                if (_cachedUnsubmittedMeetings != null && DateTime.Now < _cacheExpiry)
                {
                    System.Diagnostics.Debug.WriteLine("Using cached unsubmitted meetings");
                    unsubmittedMeetings = _cachedUnsubmittedMeetings;
                }
                else
                {
                    unsubmittedMeetings = await GetUnsubmittedMeetingsFromOutlookAsync(email);
                    _cachedUnsubmittedMeetings = unsubmittedMeetings;
                    _cacheExpiry = DateTime.Now.AddMinutes(2);
                }

                // Group by date categories
                var today = DateTime.Today;
                var yesterday = today.AddDays(-1);
                var lastWeekStart = today.AddDays(-7); // matches the fetch window (midnight, 7 days ago)

                var todayMeetings = unsubmittedMeetings.Where(m => m.StartTorontoTime.Date == today).OrderBy(m => m.StartTorontoTime).ToList();
                var yesterdayMeetings = unsubmittedMeetings.Where(m => m.StartTorontoTime.Date == yesterday).OrderBy(m => m.StartTorontoTime).ToList();
                var lastWeekMeetings = unsubmittedMeetings.Where(m => m.StartTorontoTime.Date >= lastWeekStart && m.StartTorontoTime.Date < yesterday).OrderBy(m => m.StartTorontoTime).ToList();

                // ✅ CRITICAL FIX: ALL UI updates must be on UI thread
                this.Invoke((MethodInvoker)delegate
                {
                    try
                    {
                        DisposeAndClearControls(flowUnsubmitted.Controls);

                        if (unsubmittedMeetings.Count == 0)
                        {
                            flowUnsubmitted.Controls.Add(new Label
                            {
                                Text = "No unsubmitted events found.",
                                Font = _fontLabel,
                                Size = new Size(GetFlowContentWidth(flowUnsubmitted), 40),
                                ForeColor = Color.Green,
                                Padding = new Padding(0, 10, 0, 0)
                            });
                            return;
                        }

                        if (todayMeetings.Count > 0) AddMeetingSection("Today", todayMeetings);
                        if (yesterdayMeetings.Count > 0) AddMeetingSection("Yesterday", yesterdayMeetings);
                        if (lastWeekMeetings.Count > 0) AddMeetingSection("Last Week", lastWeekMeetings);

                        // Bottom padding spacer to ensure last section is fully visible when scrolling
                        flowUnsubmitted.Controls.Add(new Label
                        {
                            Size = new Size(GetFlowContentWidth(flowUnsubmitted), 100),
                            Text = ""
                        });
                    }
                    catch (Exception invokeEx)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error during Invoke: {invokeEx.Message}");
                        System.Diagnostics.Debug.WriteLine($"Stack trace: {invokeEx.StackTrace}");
                        throw;
                    }
                });
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"LoadUnsubmittedMeetingsAsync failed: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                ShowErrorMessage($"Error loading events: {ex.Message}");
            }
        }

        // Shared card layout for both submitted and unsubmitted sections.
        // Submitted adds one extra "Program | Activity" line between the time and buttons,
        // but every other measurement (fonts, colors, gaps, button style) is identical.
        private void AddMeetingSection(string title, List<MeetingRecord> meetings, bool isSubmitted = false)
        {
            var flowPanel = isSubmitted ? flowSubmitted : flowUnsubmitted;
            int w = GetFlowContentWidth(flowPanel);

            // ── shared constants (match unsubmitted exactly) ──────────────────
            const int subjectY    = 5;
            const int subjectH    = 24;
            const int timeY       = 30;
            const int timeH       = 20;
            const int programY    = 52;   // only used when isSubmitted
            const int programH    = 18;
            const int btnY_base   = 53;   // unsubmitted: buttons start here
            const int btnH        = 27;
            const int btnGap      = 5;    // gap between the two buttons

            // submitted buttons sit one row lower (time → program → buttons)
            int btnY        = isSubmitted ? programY + programH + 3 : btnY_base;
            // submitted gets more bottom clearance: border + DPI rounding can clip the last row
            int panelHeight = btnY + btnH + (isSubmitted ? 16 : 6);

            // header colour differs so the two tabs are visually distinct;
            // everything else (panel bg, fonts, sizes) is identical
            var headerColor = isSubmitted ? Color.FromArgb(0, 150, 0) : Color.FromArgb(0, 120, 212);

            flowPanel.Controls.Add(new Label
            {
                Text      = $"{title} ({meetings.Count})",
                Font      = _fontSectionHeader,
                Size      = new Size(w, 30),
                Margin    = new Padding(0, 5, 0, 0),
                ForeColor = headerColor,
                TextAlign = ContentAlignment.MiddleLeft
            });

            foreach (var m in meetings)
            {
                var p = new Panel
                {
                    Size        = new Size(w, panelHeight),
                    BorderStyle = BorderStyle.FixedSingle,
                    BackColor   = Color.FromArgb(250, 250, 250),
                    Margin      = new Padding(0, 0, 0, 6)
                };

                // Subject
                p.Controls.Add(new Label
                {
                    Text        = m.Subject,
                    Font        = _fontSubject,
                    Location    = new Point(5, subjectY),
                    Size        = new Size(w - 15, subjectH),
                    AutoEllipsis = true
                });

                // Date / duration
                p.Controls.Add(new Label
                {
                    Text      = $"{m.StartTorontoTime:MMM dd, HH:mm} ({(m.EndUtc - m.StartUtc).TotalHours:F1} hrs)",
                    Font      = _fontTime,
                    Location  = new Point(5, timeY),
                    Size      = new Size(w - 15, timeH),
                    ForeColor = Color.FromArgb(100, 100, 100)
                });

                // Program info (submitted only) — same font/colour as the time label
                if (isSubmitted)
                {
                    p.Controls.Add(new Label
                    {
                        Text         = $"Program: {m.ProgramCode} | Activity: {m.ActivityCode}",
                        Font         = _fontTime,
                        Location     = new Point(5, programY),
                        Size         = new Size(w - 15, programH),
                        ForeColor    = Color.FromArgb(100, 100, 100),
                        AutoEllipsis = true
                    });
                }

                // Buttons — same size, same style, same gap
                int bw = (w - 15 - btnGap) / 2;

                if (isSubmitted)
                {
                    var btnCancel = new Button
                    {
                        Text      = "Cancel Submit",
                        Location  = new Point(5, btnY),
                        Size      = new Size(bw, btnH),
                        BackColor = Color.FromArgb(220, 100, 100),
                        ForeColor = Color.White,
                        FlatStyle = FlatStyle.Flat,
                        Font      = _fontButtonBold
                    };
                    btnCancel.FlatAppearance.BorderSize = 0;
                    var recordsToCancel = meetings;
                    btnCancel.Click += async (s, e) => await CancelSubmissionAsync(recordsToCancel);

                    var btnUnignore = new Button
                    {
                        Text      = "Un-Ignore",
                        Location  = new Point(5 + bw + btnGap, btnY),
                        Size      = new Size(bw, btnH),
                        BackColor = Color.FromArgb(220, 220, 220),
                        ForeColor = Color.FromArgb(60, 60, 60),
                        FlatStyle = FlatStyle.Flat,
                        Font      = _fontButton
                    };
                    btnUnignore.FlatAppearance.BorderSize = 0;
                    btnUnignore.Click += async (s, e) => await CancelIgnoreSubmissionAsync(m);

                    p.Controls.AddRange(new Control[] { btnCancel, btnUnignore });
                }
                else
                {
                    var btnSubmit = new Button
                    {
                        Text      = "Submit",
                        Location  = new Point(5, btnY),
                        Size      = new Size(bw, btnH),
                        BackColor = Color.FromArgb(0, 120, 212),
                        ForeColor = Color.White,
                        FlatStyle = FlatStyle.Flat,
                        Font      = _fontButtonBold
                    };
                    btnSubmit.FlatAppearance.BorderSize = 0;
                    btnSubmit.Click += async (s, e) => await SubmitMeetingAsync(m);

                    var btnIgnore = new Button
                    {
                        Text      = "Ignore",
                        Location  = new Point(5 + bw + btnGap, btnY),
                        Size      = new Size(bw, btnH),
                        BackColor = Color.FromArgb(220, 220, 220),
                        ForeColor = Color.FromArgb(60, 60, 60),
                        FlatStyle = FlatStyle.Flat,
                        Font      = _fontButton
                    };
                    btnIgnore.FlatAppearance.BorderSize = 0;
                    btnIgnore.Click += (s, e) => IgnoreMeeting(m, p);

                    p.Controls.AddRange(new Control[] { btnSubmit, btnIgnore });
                }

                flowPanel.Controls.Add(p);
            }

            flowPanel.Controls.Add(new Label { Size = new Size(w, 10) });
        }

        // ✅ Helper to collect all recipients from a meeting
        private static string GetAllRecipients(Outlook.AppointmentItem appt)
        {
            if (appt == null) return string.Empty;
            if (appt.MeetingStatus == Outlook.OlMeetingStatus.olNonMeeting)
                return string.Empty;

            var recipientEmails = new List<string>();
            Outlook.Recipients recipients = null;
            try
            {
                recipients = appt.Recipients;
                if (recipients != null && recipients.Count > 0)
                {
                    foreach (Outlook.Recipient recipient in recipients)
                    {
                        Outlook.Recipient r = recipient;
                        try
                        {
                            var email = GetRecipientEmail(r);
                            if (!string.IsNullOrWhiteSpace(email))
                                recipientEmails.Add(email);
                        }
                        catch { }
                        finally
                        {
                            if (r != null) Marshal.ReleaseComObject(r);
                        }
                    }
                }
                var organizerEmail = GetRecipientEmailFromAppointment(appt);
                if (!string.IsNullOrWhiteSpace(organizerEmail) && !recipientEmails.Contains(organizerEmail))
                    recipientEmails.Insert(0, organizerEmail);
            }
            catch { }
            finally
            {
                if (recipients != null) { Marshal.ReleaseComObject(recipients); }
            }
            return string.Join("; ", recipientEmails);
        }

        private static string GetRecipientEmail(Outlook.Recipient recipient)
        {
            if (recipient == null) return string.Empty;
            Outlook.AddressEntry addrEntry = null;
            try
            {
                addrEntry = recipient.AddressEntry;
                if (addrEntry != null)
                {
                    if ("EX".Equals(addrEntry.Type, StringComparison.OrdinalIgnoreCase))
                    {
                        const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
                        Outlook.PropertyAccessor pa = null;
                        try
                        {
                            pa = addrEntry.PropertyAccessor;
                            var smtp = pa.GetProperty(PR_SMTP_ADDRESS) as string;
                            if (!string.IsNullOrWhiteSpace(smtp)) return smtp;
                        }
                        finally { if (pa != null) Marshal.ReleaseComObject(pa); }
                    }
                    if (!string.IsNullOrWhiteSpace(addrEntry.Address))
                        return addrEntry.Address;
                }
                return recipient.Address ?? string.Empty;
            }
            catch { return string.Empty; }
            finally
            {
                if (addrEntry != null) { Marshal.ReleaseComObject(addrEntry); }
            }
        }

        private static string GetRecipientEmailFromAppointment(Outlook.AppointmentItem appt)
        {
            if (appt == null) return string.Empty;
            Outlook.AddressEntry organizerEntry = null;
            Outlook.NameSpace session = null;
            Outlook.AddressEntry addrEntry = null;
            Outlook.Recipient currentUser = null;
            Outlook.AddressEntry currentUserAddrEntry = null;

            try
            {
                organizerEntry = appt.GetOrganizer();
                if (organizerEntry != null)
                {
                    session = Globals.ThisAddIn?.Application?.Session;
                    if (session != null)
                    {
                        addrEntry = session.GetAddressEntryFromID(organizerEntry.ID);
                        if (addrEntry != null)
                        {
                            if ("EX".Equals(addrEntry.Type, StringComparison.OrdinalIgnoreCase))
                            {
                                const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
                                Outlook.PropertyAccessor pa = null;
                                try
                                {
                                    pa = addrEntry.PropertyAccessor;
                                    var smtp = pa.GetProperty(PR_SMTP_ADDRESS) as string;
                                    if (!string.IsNullOrWhiteSpace(smtp)) return smtp;
                                }
                                finally { if (pa != null) Marshal.ReleaseComObject(pa); }
                            }
                            if (!string.IsNullOrWhiteSpace(addrEntry.Address))
                                return addrEntry.Address;
                        }
                    }
                }
            }
            catch { }
            finally
            {
                if (addrEntry != null) { Marshal.ReleaseComObject(addrEntry); addrEntry = null; }
                if (organizerEntry != null) { Marshal.ReleaseComObject(organizerEntry); }
                if (session != null) { Marshal.ReleaseComObject(session); session = null; }
            }

            try
            {
                session = Globals.ThisAddIn?.Application?.Session;
                currentUser = session?.CurrentUser;
                if (currentUser != null)
                {
                    currentUserAddrEntry = currentUser.AddressEntry;
                    if (currentUserAddrEntry != null)
                    {
                        if ("EX".Equals(currentUserAddrEntry.Type, StringComparison.OrdinalIgnoreCase))
                        {
                            const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
                            Outlook.PropertyAccessor pa = null;
                            try
                            {
                                pa = currentUserAddrEntry.PropertyAccessor;
                                var smtp = pa.GetProperty(PR_SMTP_ADDRESS) as string;
                                if (!string.IsNullOrWhiteSpace(smtp)) return smtp;
                            }
                            finally { if (pa != null) Marshal.ReleaseComObject(pa); }
                        }
                        if (!string.IsNullOrWhiteSpace(currentUserAddrEntry.Address))
                            return currentUserAddrEntry.Address;
                    }
                }
            }
            catch { }
            finally
            {
                if (currentUserAddrEntry != null) { Marshal.ReleaseComObject(currentUserAddrEntry); }
                if (currentUser != null) { Marshal.ReleaseComObject(currentUser); }
                if (session != null) { Marshal.ReleaseComObject(session); }
            }
            return string.Empty;
        }

        // === Full Program Picker Dialog with Multi-Program Support ===
        private class ProgramPickerForm : Form
        {
            private ComboBox cboProgram;
            private ComboBox cboActivity;
            private ComboBox cboStage;
            private CheckBox chkMultiplePrograms;
            private Panel pnlMultiProgram;
            private FlowLayoutPanel flowPrograms;
            private Button btnAddProgram;
            private Label lblTotalTime;
            private Label lblAllocatedTime;
            private Button btnOk, btnCancel;

            private Font _fontAllocatedTime;

            private double _meetingDurationHours;
            private List<ProgramAllocation> _programAllocations = new List<ProgramAllocation>();

            private List<StageCodeData> _stageCodes = new List<StageCodeData>();
            private List<ActivityCodeData> _activityCodes = new List<ActivityCodeData>();
            private List<string> _programCodes = new List<string>();

            public bool IsMultiProgram => chkMultiplePrograms?.Checked ?? false;
            public List<ProgramAllocation> ProgramAllocations => _programAllocations;

            public string ProgramCode => cboProgram.SelectedItem?.ToString() ?? string.Empty;
            public string ActivityCode => GetActivityCode(cboActivity.SelectedItem);
            public string StageCode => GetStageCode(cboStage.SelectedItem);

            private string GetActivityCode(object selectedItem)
            {
                if (selectedItem is ActivityCodeData actData)
                    return actData.ActivityCode;
                return selectedItem?.ToString() ?? string.Empty;
            }

            private string GetStageCode(object selectedItem)
            {
                if (selectedItem is StageCodeData stageData)
                    return stageData.StageCode;
                return selectedItem?.ToString() ?? string.Empty;
            }


            private async Task LoadActivitiesFromSqlAsync(string initActivity)
            {
                try
                {
                    var cs = ConfigurationManager.ConnectionStrings["OemsDatabase"].ConnectionString;
                    var list = new List<ActivityCodeData>();
                    using (var cn = new SqlConnection(cs))
                    using (var cmd = new SqlCommand("dbo.Timesheet_GetActivityCodes", cn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        await cn.OpenAsync();
                        using (var rdr = await cmd.ExecuteReaderAsync())
                        {
                            while (await rdr.ReadAsync())
                            {
                                list.Add(new ActivityCodeData
                                {
                                    ActivityCode = rdr["ActivityCode"] as string ?? "",
                                    ActivityDescription = rdr["ActivityDescription"] as string ?? "",
                                    SortOrder = rdr["SortOrder"] != DBNull.Value ? Convert.ToInt32(rdr["SortOrder"]) : 0
                                });
                            }
                        }
                    }
                    _activityCodes = list;
                    cboActivity.BeginUpdate();
                    try
                    {
                        cboActivity.Items.Clear();
                        if (list.Count == 0)
                            cboActivity.Items.AddRange(new object[] { "Air", "Accommodation", "Food and Beverage", "Side Excursions", "OTHER" });
                        else
                        {
                            cboActivity.Items.AddRange(list.ToArray());
                            if (!string.IsNullOrWhiteSpace(initActivity))
                            {
                                var match = list.FirstOrDefault(a => a.ActivityCode.Equals(initActivity, StringComparison.OrdinalIgnoreCase));
                                if (match != null) cboActivity.SelectedItem = match;
                            }
                        }
                        if (cboActivity.SelectedIndex < 0 && cboActivity.Items.Count > 0) cboActivity.SelectedIndex = 0;
                        cboActivity.Enabled = true;
                    }
                    finally { cboActivity.EndUpdate(); }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"LoadActivitiesFromSqlAsync failed: {ex.Message}");
                    cboActivity.BeginUpdate();
                    try
                    {
                        cboActivity.Items.Clear();
                        cboActivity.Items.AddRange(new object[] { "Air", "Accommodation", "Food and Beverage", "Side Excursions", "OTHER" });
                        if (cboActivity.SelectedIndex < 0) cboActivity.SelectedIndex = 0;
                        cboActivity.Enabled = true;
                    }
                    finally { cboActivity.EndUpdate(); }
                }
            }

            private async Task LoadStagesFromSqlAsync(string initStage)
            {
                try
                {
                    var cs = ConfigurationManager.ConnectionStrings["OemsDatabase"].ConnectionString;
                    var list = new List<StageCodeData>();
                    using (var cn = new SqlConnection(cs))
                    using (var cmd = new SqlCommand("dbo.Timesheet_GetStageCodes", cn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        var userEmail = GetCurrentUserEmail();
                        cmd.Parameters.Add(new SqlParameter("@UserEmail", SqlDbType.NVarChar, 320) { Value = userEmail });
                        await cn.OpenAsync();
                        using (var rdr = await cmd.ExecuteReaderAsync())
                        {
                            while (await rdr.ReadAsync())
                            {
                                list.Add(new StageCodeData
                                {
                                    StageCode = rdr["StageCode"] as string ?? "",
                                    StageDescription = rdr["StageDescription"] as string ?? "",
                                    SortOrder = rdr["SortOrder"] != DBNull.Value ? Convert.ToInt32(rdr["SortOrder"]) : 0
                                });
                            }
                        }
                    }
                    _stageCodes = list;
                    cboStage.BeginUpdate();
                    try
                    {
                        cboStage.Items.Clear();
                        if (list.Count == 0)
                            cboStage.Items.AddRange(new object[] { "Client Communication", "Internal Communication", "Vendor Communication", "Work Time" });
                        else
                        {
                            cboStage.Items.AddRange(list.ToArray());
                            if (!string.IsNullOrWhiteSpace(initStage))
                            {
                                var match = list.FirstOrDefault(s => s.StageCode.Equals(initStage, StringComparison.OrdinalIgnoreCase));
                                if (match != null) cboStage.SelectedItem = match;
                            }
                        }
                        if (cboStage.SelectedIndex < 0 && cboStage.Items.Count > 0) cboStage.SelectedIndex = 0;
                        cboStage.Enabled = true;
                    }
                    finally { cboStage.EndUpdate(); }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"LoadStagesFromSqlAsync failed: {ex.Message}");
                    cboStage.BeginUpdate();
                    try
                    {
                        cboStage.Items.Clear();
                        cboStage.Items.AddRange(new object[] { "Client Communication", "Internal Communication", "Vendor Communication", "Work Time" });
                        if (cboStage.SelectedIndex < 0) cboStage.SelectedIndex = 0;
                        cboStage.Enabled = true;
                    }
                    finally { cboStage.EndUpdate(); }
                }
            }

            private async Task LoadProgramsFromSqlAsync(string email, string initProgram)
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(email)) throw new ArgumentException("Email required");
                    var cs = ConfigurationManager.ConnectionStrings["OemsDatabase"].ConnectionString;
                    var list = new List<string>();
                    using (var cn = new SqlConnection(cs))
                    {
                        await cn.OpenAsync();
                        using (var cmd = new SqlCommand("dbo.TimeSheet_GetActivePrograms", cn))
                        {
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.CommandTimeout = 30;
                            cmd.Parameters.Add(new SqlParameter("@UserEmail", SqlDbType.NVarChar, 320) { Value = email });
                            using (var rdr = await cmd.ExecuteReaderAsync())
                            {
                                while (await rdr.ReadAsync())
                                {
                                    var code = rdr["ProgramCode"] as string;
                                    if (!string.IsNullOrWhiteSpace(code))
                                        list.Add(code.Trim());
                                }
                            }
                        }
                    }
                    _programCodes = list;
                    cboProgram.BeginUpdate();
                    try
                    {
                        cboProgram.Items.Clear();
                        cboProgram.Items.Add("Project-OEMS A12004");
                        cboProgram.Items.Add("Project Monday A12005");
                        cboProgram.Items.Add("Buying/Proposal A13000");
                        cboProgram.Items.Add("People-Vacation A14001");
                        cboProgram.Items.Add("People-Personal Time A14002");
                        cboProgram.Items.Add("People-Sick A14003");
                        cboProgram.Items.Add("People-Stat Holiday A14004");
                        cboProgram.Items.Add("Finance-Invoicing/AR A15001");
                        if (list.Count > 0) cboProgram.Items.Add("──────────────────────");
                        if (list.Count == 0)
                            cboProgram.Items.Add("(No active programs found)");
                        else
                            cboProgram.Items.AddRange(list.ToArray());
                        SelectIfPresent(cboProgram, initProgram);
                        if (cboProgram.SelectedIndex < 0) cboProgram.SelectedIndex = 0;
                        cboProgram.Enabled = true;
                    }
                    finally { cboProgram.EndUpdate(); }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"LoadProgramsFromSqlAsync failed: {ex.Message}");
                    cboProgram.BeginUpdate();
                    try
                    {
                        cboProgram.Items.Clear();
                        cboProgram.Items.Add("Project-OEMS A12004");
                        cboProgram.Items.Add("Project Monday A12005");
                        cboProgram.Items.Add("Buying/Proposal A13000");
                        cboProgram.Items.Add("People-Vacation A14001");
                        cboProgram.Items.Add("People-Personal Time A14002");
                        cboProgram.Items.Add("People-Sick A14003");
                        cboProgram.Items.Add("People-Stat Holiday A14004");
                        cboProgram.Items.Add("Finance-Invoicing/AR A15001");
                        cboProgram.Items.Add("──────────────────────");
                        cboProgram.Items.AddRange(new object[] { "test-0001", "test-0002" });
                        if (cboProgram.SelectedIndex < 0) cboProgram.SelectedIndex = 0;
                        cboProgram.Enabled = true;
                    }
                    finally { cboProgram.EndUpdate(); }
                    MessageBox.Show("Failed to load Program/Proposal codes: " + ex.Message);
                }
            }

            public ProgramPickerForm() : this(null, null, null, 0) { }
            public ProgramPickerForm(string initProgram, string initActivity, string initStage) : this(initProgram, initActivity, initStage, 0) { }

            public ProgramPickerForm(string initProgram, string initActivity, string initStage, double meetingDurationHours)
            {
                _meetingDurationHours = meetingDurationHours;

                Text = "Meeting Information";
                FormBorderStyle = FormBorderStyle.FixedDialog;
                StartPosition = FormStartPosition.CenterScreen;
                MaximizeBox = false;
                MinimizeBox = false;
                Width = 520;
                Height = 220;
                TopMost = true;

                var lblProgram = new Label { Left = 15, Top = 20, Width = 120, Text = "Program Code:" };
                cboProgram = new ComboBox { Left = 140, Top = 16, Width = 340, DropDownStyle = ComboBoxStyle.DropDownList, TabIndex = 1 };

                var lblActivity = new Label { Left = 15, Top = 55, Width = 120, Text = "Activity Code:" };
                cboActivity = new ComboBox { Left = 140, Top = 51, Width = 340, DropDownStyle = ComboBoxStyle.DropDownList, TabIndex = 2 };

                var lblStage = new Label { Left = 15, Top = 90, Width = 120, Text = "Stage Code:" };
                cboStage = new ComboBox { Left = 140, Top = 86, Width = 340, DropDownStyle = ComboBoxStyle.DropDownList, TabIndex = 3 };

                chkMultiplePrograms = new CheckBox
                {
                    Left = 15, Top = 125, Width = 290,
                    Text = "Add Additional Programs",
                    TabIndex = 4,
                    Visible = meetingDurationHours > 0
                };
                chkMultiplePrograms.CheckedChanged += ChkMultiplePrograms_CheckedChanged;

                pnlMultiProgram = new Panel
                {
                    Left = 15, Top = 155, Width = 475, Height = 270,
                    BorderStyle = BorderStyle.FixedSingle,
                    Visible = false,
                    BackColor = Color.FromArgb(250, 250, 250)
                };

                flowPrograms = new FlowLayoutPanel
                {
                    Left = 10, Top = 30, Width = 450, Height = 200,
                    AutoScroll = true,
                    FlowDirection = FlowDirection.TopDown,
                    WrapContents = false,
                    BorderStyle = BorderStyle.None
                };

                btnAddProgram = new Button
                {
                    Left = 10, Top = 235, Width = 150, Height = 25,
                    Text = "+ Add Program",
                    BackColor = Color.FromArgb(0, 120, 212),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat
                };
                btnAddProgram.FlatAppearance.BorderSize = 0;
                btnAddProgram.Click += BtnAddProgram_Click;

                _fontAllocatedTime = new Font("Segoe UI", 9, FontStyle.Bold);
                lblAllocatedTime = new Label
                {
                    Left = 170, Top = 238, Width = 295, Height = 25,
                    Text = $"Allocated: {_meetingDurationHours:F1} / {_meetingDurationHours:F1} hrs",
                    Font = _fontAllocatedTime,
                    ForeColor = Color.Green,
                    TextAlign = ContentAlignment.MiddleRight
                };

                pnlMultiProgram.Controls.AddRange(new Control[] { flowPrograms, btnAddProgram, lblAllocatedTime });

                btnOk = new Button { Left = 310, Top = 123, Width = 80, Height = 28, Text = "OK", DialogResult = DialogResult.OK, TabIndex = 5 };
                btnCancel = new Button { Left = 400, Top = 123, Width = 80, Height = 28, Text = "Cancel", DialogResult = DialogResult.Cancel, TabIndex = 6 };
                btnOk.Click += BtnOk_Click;

                Controls.AddRange(new Control[] {
                    lblProgram, cboProgram, lblActivity, cboActivity, lblStage, cboStage,
                    chkMultiplePrograms, pnlMultiProgram, btnOk, btnCancel
                });
                AcceptButton = btnOk;
                CancelButton = btnCancel;

                cboProgram.Items.Add("Loading..."); cboProgram.SelectedIndex = 0; cboProgram.Enabled = false;
                cboActivity.Items.Add("Loading..."); cboActivity.SelectedIndex = 0; cboActivity.Enabled = false;
                cboStage.Items.Add("Loading..."); cboStage.SelectedIndex = 0; cboStage.Enabled = false;

                var programToSelect = initProgram;
                var activityToSelect = initActivity;
                var stageToSelect = initStage;

                this.Shown += async (s, e) =>
                {
                    var user = GetCurrentUserEmail() ?? string.Empty;
                    await Task.WhenAll(
                        LoadProgramsFromSqlAsync(user, programToSelect),
                        LoadActivitiesFromSqlAsync(activityToSelect),
                        LoadStagesFromSqlAsync(stageToSelect)
                    );
                };
            }

            private void ChkMultiplePrograms_CheckedChanged(object sender, EventArgs e)
            {
                bool isMulti = chkMultiplePrograms.Checked;
                if (isMulti)
                {
                    this.Height = 530;
                    btnOk.Top = 435; btnCancel.Top = 435;
                    btnOk.Left = 320; btnCancel.Left = 410;
                    cboProgram.Enabled = true; cboActivity.Enabled = true; cboStage.Enabled = true;
                    pnlMultiProgram.Visible = true;
                }
                else
                {
                    this.Height = 220;
                    btnOk.Top = 123; btnCancel.Top = 123;
                    btnOk.Left = 310; btnCancel.Left = 400;
                    cboProgram.Enabled = true; cboActivity.Enabled = true; cboStage.Enabled = true;
                    pnlMultiProgram.Visible = false;
                    foreach (Control ctrl in flowPrograms.Controls.OfType<ProgramAllocationControl>().ToList())
                    {
                        flowPrograms.Controls.Remove(ctrl);
                        ctrl.Dispose();
                    }
                    _programAllocations.Clear();
                }
            }

            private void BtnAddProgram_Click(object sender, EventArgs e) => AddProgramAllocationControl();

            private void AddProgramAllocationControl(string initProgram = null, string initActivity = null, string initStage = null)
            {
                var allocationControl = new ProgramAllocationControl(
                    _meetingDurationHours,
                    cboProgram.Items.Cast<string>().ToList(),
                    initProgram, initActivity, initStage
                );
                allocationControl.OnRemove += (s, ev) =>
                {
                    var ctrl = s as ProgramAllocationControl;
                    _programAllocations.Remove(ctrl.Allocation);
                    flowPrograms.Controls.Remove(ctrl);
                    ctrl.Dispose();
                    UpdateTotalAllocated();
                };
                allocationControl.OnHoursChanged += (s, ev) => UpdateTotalAllocated();
                flowPrograms.Controls.Add(allocationControl);
                _programAllocations.Add(allocationControl.Allocation);
                UpdateTotalAllocated();
            }

            private void UpdateTotalAllocated()
            {
                double additionalHours = _programAllocations.Sum(p => p.Hours);
                double originalProgramHours = _meetingDurationHours - additionalHours;
                string originalProgram = cboProgram.SelectedItem?.ToString() ?? "Original";
                lblAllocatedTime.Text = $"Total: {_meetingDurationHours:F1}h ({originalProgram}: {originalProgramHours:F1}h)";
                bool isValid = originalProgramHours > 0.01;
                lblAllocatedTime.ForeColor = isValid ? Color.Green : Color.Red;
            }

            private void BtnOk_Click(object sender, EventArgs e)
            {
                if (string.IsNullOrWhiteSpace(cboProgram.SelectedItem?.ToString()))
                {
                    MessageBox.Show("Please select a program code.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    DialogResult = DialogResult.None;
                    return;
                }

                if (chkMultiplePrograms != null && chkMultiplePrograms.Checked)
                {
                    if (_programAllocations.Count < 1)
                    {
                        MessageBox.Show(
                            "Please add at least 1 additional program.\n\nIf this meeting only involves one program, please uncheck the 'Add Additional Programs' checkbox.",
                            "Additional Programs Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        DialogResult = DialogResult.None;
                        return;
                    }

                    string originalProgram = cboProgram.SelectedItem?.ToString() ?? "";
                    var duplicatePrograms = _programAllocations
                        .Where(p => !string.IsNullOrWhiteSpace(p.ProgramCode) &&
                                   p.ProgramCode.Equals(originalProgram, StringComparison.OrdinalIgnoreCase))
                        .ToList();

                    if (duplicatePrograms.Count > 0)
                    {
                        MessageBox.Show(
                            $"Cannot add '{originalProgram}' as an additional program because it's already the original program.\n\nPlease select a different program for the additional allocation.",
                            "Duplicate Program Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        DialogResult = DialogResult.None;
                        return;
                    }

                    double additionalHours = _programAllocations.Sum(p => p.Hours);
                    if (additionalHours >= _meetingDurationHours)
                    {
                        MessageBox.Show(
                            $"Additional programs ({additionalHours:F2} hrs) cannot equal or exceed total meeting duration ({_meetingDurationHours:F2} hrs).\n\nPlease leave time for the original program.",
                            "Invalid Allocation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        DialogResult = DialogResult.None;
                        return;
                    }

                    if (_programAllocations.Any(p => string.IsNullOrWhiteSpace(p.ProgramCode)))
                    {
                        MessageBox.Show("Please select a program code for all additional entries.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        DialogResult = DialogResult.None;
                        return;
                    }
                }
            }

            private static void SelectIfPresent(ComboBox combo, string value)
            {
                if (string.IsNullOrWhiteSpace(value)) return;
                var idx = combo.FindStringExact(value);
                if (idx >= 0) combo.SelectedIndex = idx;
            }

            private static string GetCurrentUserEmail()
            {
                Outlook.NameSpace session = null;
                Outlook.Recipient currentUser = null;
                Outlook.AddressEntry addrEntry = null;
                try
                {
                    session = Globals.ThisAddIn.Application.Session;
                    currentUser = session.CurrentUser;
                    addrEntry = currentUser?.AddressEntry;
                    if (addrEntry != null)
                    {
                        if ("EX".Equals(addrEntry.Type, StringComparison.OrdinalIgnoreCase))
                        {
                            const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
                            Outlook.PropertyAccessor pa = null;
                            try
                            {
                                pa = addrEntry.PropertyAccessor;
                                var smtp = pa.GetProperty(PR_SMTP_ADDRESS) as string;
                                if (!string.IsNullOrWhiteSpace(smtp)) return smtp;
                            }
                            finally { if (pa != null) Marshal.ReleaseComObject(pa); }
                        }
                        if (!string.IsNullOrWhiteSpace(addrEntry.Address)) return addrEntry.Address;
                    }
                    return currentUser?.Name ?? string.Empty;
                }
                catch { return string.Empty; }
                finally
                {
                    if (addrEntry != null) { Marshal.ReleaseComObject(addrEntry); addrEntry = null; }
                    if (currentUser != null) { Marshal.ReleaseComObject(currentUser); currentUser = null; }
                    if (session != null) { Marshal.ReleaseComObject(session); session = null; }
                }
            }

            protected override void Dispose(bool disposing)
            {
                if (disposing) _fontAllocatedTime?.Dispose();
                base.Dispose(disposing);
            }
        }

        private class StageCodeData
        {
            public string StageCode { get; set; }
            public string StageDescription { get; set; }
            public int SortOrder { get; set; }
            public override string ToString() => StageDescription;
        }

        private class ActivityCodeData
        {
            public string ActivityCode { get; set; }
            public string ActivityDescription { get; set; }
            public int SortOrder { get; set; }
            public override string ToString() => ActivityDescription;
        }

        private class ProgramAllocationControl : Panel
        {
            public ProgramAllocation Allocation { get; private set; }
            public event EventHandler OnRemove;
            public event EventHandler OnHoursChanged;

            private ComboBox cboProgram;
            private ComboBox cboActivity;
            private ComboBox cboStage;
            private TrackBar trackHours;
            private Label lblHours;
            private Button btnRemove;
            private Font _fontHoursLabel;
            private Font _fontRemoveButton;
            private double _maxHours;

            public ProgramAllocationControl(double maxHours, List<string> programs, string initProgram = null, string initActivity = null, string initStage = null)
            {
                _maxHours = maxHours;
                Allocation = new ProgramAllocation { Hours = maxHours };

                Width = 420; Height = 135;
                BorderStyle = BorderStyle.FixedSingle;
                Margin = new Padding(0, 0, 0, 8);
                BackColor = Color.FromArgb(250, 250, 250);

                cboProgram = new ComboBox { Left = 10, Top = 10, Width = 330, DropDownStyle = ComboBoxStyle.DropDownList };
                cboProgram.Items.AddRange(programs.ToArray());
                if (!string.IsNullOrWhiteSpace(initProgram))
                {
                    var idx = cboProgram.FindStringExact(initProgram);
                    if (idx >= 0) cboProgram.SelectedIndex = idx;
                }
                if (cboProgram.SelectedIndex < 0 && cboProgram.Items.Count > 0) cboProgram.SelectedIndex = 0;
                cboProgram.SelectedIndexChanged += (s, e) => Allocation.ProgramCode = cboProgram.SelectedItem?.ToString() ?? "";

                cboActivity = new ComboBox { Left = 10, Top = 40, Width = 330, DropDownStyle = ComboBoxStyle.DropDownList };
                cboActivity.Items.AddRange(new object[] { "Work Time", "Client Communication", "Vendor Communication", "Internal Communication" });
                if (!string.IsNullOrWhiteSpace(initActivity))
                {
                    var idx = cboActivity.FindStringExact(initActivity);
                    if (idx >= 0) cboActivity.SelectedIndex = idx;
                }
                if (cboActivity.SelectedIndex < 0 && cboActivity.Items.Count > 0) cboActivity.SelectedIndex = 0;
                cboActivity.SelectedIndexChanged += (s, e) => Allocation.ActivityCode = cboActivity.SelectedItem?.ToString() ?? "";

                cboStage = new ComboBox { Left = 10, Top = 70, Width = 330, DropDownStyle = ComboBoxStyle.DropDownList };
                cboStage.Items.AddRange(new object[] { "Client Meeting", "Internal Meeting", "Email", "Vendor Research", "Meeting", "Timesheet", "Design", "Registration" });
                if (!string.IsNullOrWhiteSpace(initStage))
                {
                    var idx = cboStage.FindStringExact(initStage);
                    if (idx >= 0) cboStage.SelectedIndex = idx;
                }
                if (cboStage.SelectedIndex < 0 && cboStage.Items.Count > 0) cboStage.SelectedIndex = 0;
                cboStage.SelectedIndexChanged += (s, e) => Allocation.StageCode = cboStage.SelectedItem?.ToString() ?? "";

                int maxMinutes = (int)(maxHours * 60);
                trackHours = new TrackBar
                {
                    Left = 10, Top = 100, Width = 280,
                    Minimum = 0, Maximum = maxMinutes, Value = maxMinutes,
                    TickFrequency = 15, SmallChange = 5, LargeChange = 15
                };
                trackHours.ValueChanged += TrackHours_ValueChanged;

                int initialMinutes = (int)(Allocation.Hours * 60);
                _fontHoursLabel  = new Font("Segoe UI", 8, FontStyle.Bold);
                _fontRemoveButton = new Font("Segoe UI", 8, FontStyle.Bold);

                lblHours = new Label
                {
                    Left = 295, Top = 105, Width = 70,
                    Text = FormatTimeLabel(initialMinutes),
                    Font = _fontHoursLabel
                };

                btnRemove = new Button
                {
                    Left = 345, Top = 10, Width = 65, Height = 25,
                    Text = "Delete",
                    BackColor = Color.Red, ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat,
                    Font = _fontRemoveButton
                };
                btnRemove.FlatAppearance.BorderSize = 0;
                btnRemove.Click += (s, e) => OnRemove?.Invoke(this, EventArgs.Empty);

                Controls.AddRange(new Control[] { cboProgram, cboActivity, cboStage, trackHours, lblHours, btnRemove });

                Allocation.ProgramCode = cboProgram.SelectedItem?.ToString() ?? "";
                Allocation.ActivityCode = cboActivity.SelectedItem?.ToString() ?? "";
                Allocation.StageCode = cboStage.SelectedItem?.ToString() ?? "";
            }

            private string FormatTimeLabel(int totalMinutes)
            {
                if (totalMinutes < 60) return $"{totalMinutes}mins";
                int hours = totalMinutes / 60;
                int minutes = totalMinutes % 60;
                return minutes == 0 ? $"{hours}h" : $"{hours}h {minutes}mins";
            }

            private void TrackHours_ValueChanged(object sender, EventArgs e)
            {
                int minutes = trackHours.Value;
                Allocation.Hours = minutes / 60.0;
                lblHours.Text = FormatTimeLabel(minutes);
                OnHoursChanged?.Invoke(this, EventArgs.Empty);
            }

            protected override void Dispose(bool disposing)
            {
                if (disposing)
                {
                    _fontHoursLabel?.Dispose();
                    _fontRemoveButton?.Dispose();
                }
                base.Dispose(disposing);
            }
        }
    }
}