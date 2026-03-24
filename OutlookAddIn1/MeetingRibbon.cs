using Microsoft.Office.Tools.Ribbon;
using OutlookAddIn1;
using System;
using System.Linq;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using OutlookAddIn1.Graph;

namespace OutlookAddIn1
{
    public class MeetingRibbon : RibbonBase
    {
        private RibbonTab tab;
        private RibbonGroup group;
        private RibbonButton btnNewMeeting;

        public MeetingRibbon() : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.tab = this.Factory.CreateRibbonTab();
            this.group = this.Factory.CreateRibbonGroup();
            this.btnNewMeeting = this.Factory.CreateRibbonButton();
            this.tab.SuspendLayout();
            this.group.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab
            // 
            this.tab.Groups.Add(this.group);
            this.tab.Label = "Meetings";
            this.tab.Name = "tab";
            // 
            // group
            // 
            this.group.Items.Add(this.btnNewMeeting);
            this.group.Label = "Actions";
            this.group.Name = "group";
      
            
            this.btnNewMeeting.Label = "New Online Meeting";
            this.btnNewMeeting.Name = "btnNewMeeting";
            this.btnNewMeeting.ScreenTip = "Create a pre-filled meeting invite";
            this.btnNewMeeting.Click += button1_Click;
            // 
            // MeetingRibbon
            // 
            this.Name = "MeetingRibbon";
            this.RibbonType = "Microsoft.Outlook.Explorer, Microsoft.Outlook.Appointment";
            this.Tabs.Add(this.tab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MeetingRibbon_Load);
            this.tab.ResumeLayout(false);
            this.tab.PerformLayout();
            this.group.ResumeLayout(false);
            this.group.PerformLayout();
            this.ResumeLayout(false);

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            using (var dlg = new ProgramPickerForm())
            {
                var result = dlg.ShowDialog();
                if (result != DialogResult.OK) return;

                try
                {
                    var app = Globals.ThisAddIn.Application;
                    var appt = app.CreateItem(Outlook.OlItemType.olAppointmentItem)
                                  as Outlook.AppointmentItem;

                    // Example timings (adjust as needed)
                    appt.Subject = $"[{dlg.ProgramCode}]";
                    appt.Start = DateTime.Now.AddMinutes(30);
                    appt.End = DateTime.Now.AddMinutes(60);
                    appt.Location = "Microsoft Teams";
                    appt.MeetingStatus = Outlook.OlMeetingStatus.olMeeting;

                    appt.Body =
                        $"Program: {dlg.ProgramCode}\r\n" +
                        $"Activity: {dlg.ActivityCode}\r\n" +
                        $"Stage: {dlg.StageCode}\r\n\r\n";

                    var ups = appt.UserProperties;
                    AddOrSetTextProp(ups, "ProgramCode", dlg.ProgramCode);
                    AddOrSetTextProp(ups, "ActivityCode", dlg.ActivityCode);
                    AddOrSetTextProp(ups, "StageCode", dlg.StageCode);

                    appt.Display(false);

                    var t = new Timer { Interval = 250 };
                    t.Tick +=(s2,e2) =>
                    {
                        try
                        {
                            t.Stop();
                            PutMetaUnderTeams(appt, dlg.ProgramCode, dlg.ActivityCode, dlg.StageCode, fallbackJoinUrl: "");
                            // optional: move the cursor to top or save draft, etc.
                            // appt.Save();
                        }
                        catch { /* swallow */ }
                        finally { t.Dispose(); }
                    };
                    t.Start();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message);
                }
            }
        }

        private static void AddOrSetTextProp(Outlook.UserProperties ups, string name, string value)
        {
            var up = ups.Find(name) ?? ups.Add(name, Outlook.OlUserPropertyType.olText, false, Type.Missing);
            up.Value = value ?? string.Empty;
        }

        private class ProgramPickerForm : Form
        {
            private TextBox txtProgram;
            private ComboBox cboActivity;
            private ComboBox cboStage;
            private Button btnOk;
            private Button btnCancel;
            private Label lblProgram;
            private Label lblActivity;
            private Label lblStage;

            public string ProgramCode => txtProgram.Text?.Trim();
            public string ActivityCode => cboActivity.SelectedItem?.ToString() ?? string.Empty;
            public string StageCode => cboStage.SelectedItem?.ToString() ?? string.Empty;

            public ProgramPickerForm()
            {
                this.Text = "Meeting Information";
                this.FormBorderStyle = FormBorderStyle.FixedDialog;
                this.StartPosition = FormStartPosition.CenterParent;
                this.MaximizeBox = false;
                this.MinimizeBox = false;
                this.Width = 420;
                this.Height = 240;

                lblProgram = new Label { Left = 15, Top = 20, Width = 120, Text = "Program Code:" };
                txtProgram = new TextBox { Left = 140, Top = 16, Width = 240, TabIndex = 0 };

                lblActivity = new Label { Left = 15, Top = 65, Width = 120, Text = "Activity Code:" };
                cboActivity = new ComboBox { Left = 140, Top = 60, Width = 240, DropDownStyle = ComboBoxStyle.DropDownList, TabIndex = 1 };

                lblStage = new Label { Left = 15, Top = 105, Width = 120, Text = "Stage Code:" };
                cboStage = new ComboBox { Left = 140, Top = 100, Width = 240, DropDownStyle = ComboBoxStyle.DropDownList, TabIndex = 2 };

                btnOk = new Button { Left = 200, Top = 145, Width = 80, Text = "OK", DialogResult = DialogResult.OK, TabIndex = 3 };
                btnCancel = new Button { Left = 300, Top = 145, Width = 80, Text = "Cancel", DialogResult = DialogResult.Cancel, TabIndex = 4 };

                this.Controls.AddRange(new Control[] {
                    lblProgram, txtProgram,
                    lblActivity, cboActivity,
                    lblStage, cboStage,
                    btnOk, btnCancel
                });

                
                this.AcceptButton = btnOk;
                this.CancelButton = btnCancel;

                var programSuggestions = new[] {
                    "KIA-0623", "BLC-0618", "COU-0534"
                };

                var ac = new AutoCompleteStringCollection();
                ac.AddRange(programSuggestions);
                txtProgram.AutoCompleteCustomSource = ac;
                txtProgram.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                txtProgram.AutoCompleteSource = AutoCompleteSource.CustomSource;

                cboActivity.Items.AddRange(new object[] {
                    "Air", "Accommodation", "Food and Beverage", "Side Excursions", "OTHER"
                });
                cboStage.Items.AddRange(new object[] {
                    "Client Communication", "Internal Communication", "Vendor Communication", "Work Time"
                });

                if (cboActivity.Items.Count > 0) cboActivity.SelectedIndex = 0;
                if (cboStage.Items.Count > 0) cboStage.SelectedIndex = 0;
            }
        }

        private void MeetingRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }


        // Insert data after the invite section 
        private static void PutMetaUnderTeams(Outlook.AppointmentItem appt,
            string program, string activity, string stage, string fallbackJoinUrl = "")
        {
            string meta =
                $"Program: {program}\r\n" +
                $"Activity: {activity}\r\n" +
                $"Stage: {stage}\r\n";

            string body = appt.Body ?? string.Empty;

            int insertAt = IndexOfTeamsBlockEnd(body);
            string before = body.Substring(0, insertAt);
            string after = body.Substring(insertAt);
            appt.Body = before + "\r\n" + meta + "\r\n" + after;
            
        }

        private static int IndexOfTeamsBlockEnd(string body)
        {
            if (string.IsNullOrEmpty(body)) return -1;

            var lines = body.Split(new[] { "\r\n" }, StringSplitOptions.None);
            int charIndex = 0;

            // look for where the teams link/header appears
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i];
                bool isTeamsLine =
                    line.IndexOf("Microsoft Teams", StringComparison.OrdinalIgnoreCase) >= 0;

                // accumulate charIndex to the beginning of this line
                // (done at end of loop in first iteration; easier to recompute)
                if (isTeamsLine)
                {
                    // find the next blank line after the link/header (end of block)
                    int j = i;
                    for (; j < lines.Length; j++)
                    {
                        if (string.IsNullOrWhiteSpace(lines[j])) break;
                    }

                    // compute char offset to the end of that segment
                    int pos = 0;
                    for (int k = 0; k <= j && k < lines.Length; k++)
                        pos += lines[k].Length + 2; // +2 for CRLF

                    return pos;
                }
            }
            return -1;
        }




    }



}





