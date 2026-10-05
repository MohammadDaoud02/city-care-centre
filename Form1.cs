using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Printing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace CityCareClinic
{
    public partial class MainForm : Form
    {
        // استدعاء مكتبة Windows لإظهار النصوص التوضيحية (Placeholders)
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern Int32 SendMessage(IntPtr hWnd, int msg, int wParam, [MarshalAs(UnmanagedType.LPWStr)] string lParam);
        private const int EM_SETCUEBANNER = 0x1501;

        // دالة مساعدة لتطبيق النص التوضيحي المائي على أي TextBox
        private void SetPlaceholder(TextBox textBox, string placeholderText)
        {
            SendMessage(textBox.Handle, EM_SETCUEBANNER, 0, placeholderText);
        }

        // عناصر واجهة المستخدم الرئيسية
        private Panel sidebarPanel;
        private Panel mainContentPanel;

        // جميع الشاشات
        private Panel loginPanel;
        private Panel dashboardPanel;
        private Panel addPatientPanel;
        private Panel recordVisitPanel;
        private Panel diagnosesPanel;
        private Panel referralsPanel;
        private Panel staffPanel;
        private Panel clinicsPanel;
        private Panel reportsPanel;

        // عناصر Dashboard للبيانات الحية
        private Label lblPatientCountVal;
        private Label lblTodayVisitsVal;
        private Label lblDoctorCountVal;
        private Label lblNurseCountVal;
        private DataGridView dgvDashboardVisits;

        // عناصر الطباعة
        private PrintDocument printDoc = new PrintDocument();
        private string reportTextToPrint = "";

        public MainForm()
        {
            InitializeComponentCustom();
            ShowLoginScreen();
        }

        private void InitializeComponentCustom()
        {
            this.Text = "CityCare Clinic Management System";
            this.Size = new Size(1100, 700);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(240, 243, 246);

            // الشريط الجانبي (Sidebar)
            sidebarPanel = new Panel
            {
                Dock = DockStyle.Left,
                Width = 200,
                BackColor = Color.FromArgb(30, 41, 59),
                Visible = false
            };

            CreateSidebarButtons();

            // اللوحة الرئيسية للمحتوى
            mainContentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(245, 247, 250)
            };

            this.Controls.Add(mainContentPanel);
            this.Controls.Add(sidebarPanel);

            // تهيئة جميع الشاشات
            BuildLoginPanel();
            BuildDashboardPanel();
            BuildAddPatientPanel();
            BuildRecordVisitPanel();
            BuildDiagnosesPanel();
            BuildReferralsPanel();
            BuildStaffPanel();
            BuildClinicsPanel();
            BuildReportsPanel();
        }

        private void CreateSidebarButtons()
        {
            Label logo = new Label
            {
                Text = "CityCare Clinic",
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 14, FontStyle.Bold),
                Dock = DockStyle.Top,
                Height = 60,
                TextAlign = ContentAlignment.MiddleCenter
            };
            sidebarPanel.Controls.Add(logo);

            string[] menuItems = { "Dashboard", "Patients", "Visits", "Diagnoses", "Referrals", "Staff", "Clinics", "Reports" };
            for (int i = menuItems.Length - 1; i >= 0; i--)
            {
                Button btn = new Button
                {
                    Text = "  " + menuItems[i],
                    Dock = DockStyle.Top,
                    Height = 45,
                    FlatStyle = FlatStyle.Flat,
                    ForeColor = Color.Gainsboro,
                    Font = new Font("Segoe UI", 10),
                    TextAlign = ContentAlignment.MiddleLeft,
                    Tag = menuItems[i]
                };
                btn.FlatAppearance.BorderSize = 0;
                btn.Click += MenuButton_Click;
                sidebarPanel.Controls.Add(btn);
            }
        }

        private void MenuButton_Click(object sender, EventArgs e)
        {
            Button btn = sender as Button;
            if (btn == null) return;

            switch (btn.Tag.ToString())
            {
                case "Dashboard": 
                    ShowPanel(dashboardPanel); 
                    LoadDashboardData(); 
                    break;
                case "Patients": ShowPanel(addPatientPanel); break;
                case "Visits": ShowPanel(recordVisitPanel); break;
                case "Diagnoses": ShowPanel(diagnosesPanel); break;
                case "Referrals": ShowPanel(referralsPanel); break;
                case "Staff": ShowPanel(staffPanel); break;
                case "Clinics": ShowPanel(clinicsPanel); break;
                case "Reports": ShowPanel(reportsPanel); break;
            }
        }

        private void ShowPanel(Panel panelToShow)
        {
            loginPanel.Visible = false;
            dashboardPanel.Visible = false;
            addPatientPanel.Visible = false;
            recordVisitPanel.Visible = false;
            diagnosesPanel.Visible = false;
            referralsPanel.Visible = false;
            staffPanel.Visible = false;
            clinicsPanel.Visible = false;
            reportsPanel.Visible = false;

            panelToShow.Visible = true;
            panelToShow.BringToFront();
        }

        #region 1. Login Screen (Login by Name & Password)
        private void BuildLoginPanel()
        {
            loginPanel = new Panel { Dock = DockStyle.Fill };

            Panel card = new Panel
            {
                Size = new Size(350, 400),
                BackColor = Color.White,
                Location = new Point((1100 - 350) / 2, (700 - 400) / 2 - 20)
            };

            Label title = new Label
            {
                Text = "CityCare Clinic Management System\n\nLogin",
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                ForeColor = Color.FromArgb(15, 23, 42),
                Dock = DockStyle.Top,
                Height = 80,
                TextAlign = ContentAlignment.MiddleCenter
            };

            Label lblUser = new Label { Text = "Doctor / Nurse Name", Location = new Point(40, 100), AutoSize = true };
            TextBox txtUser = new TextBox { Location = new Point(40, 125), Width = 270, Text = "" };
            SetPlaceholder(txtUser, "Ex: Dr. Ahmed Hassan");

            Label lblPass = new Label { Text = "Password", Location = new Point(40, 165), AutoSize = true };
            TextBox txtPass = new TextBox { Location = new Point(40, 190), Width = 270, UseSystemPasswordChar = true, Text = "" };
            SetPlaceholder(txtPass, "Enter Password");

            Button btnLogin = new Button
            {
                Text = "Login",
                Location = new Point(40, 240),
                Width = 270,
                Height = 40,
                BackColor = Color.FromArgb(37, 99, 235),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };

            btnLogin.Click += (s, e) =>
            {
                string nameInput = txtUser.Text.Trim();
                string passwordInput = txtPass.Text.Trim();

                if (string.IsNullOrWhiteSpace(nameInput) || string.IsNullOrWhiteSpace(passwordInput))
                {
                    MessageBox.Show("يرجى إدخال اسم الطبيب/الممرض وكلمة المرور", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // الاستعلام والربط عن طريق Name بدلاً من EmployeeID
                string query = "SELECT EmployeeID, Name, Role FROM Staff WHERE Name = @name AND Password = @pass";
                SqlParameter[] parameters = {
            new SqlParameter("@name", nameInput),
            new SqlParameter("@pass", passwordInput)
        };

                DataTable dt = DatabaseHelper.ExecuteQuery(query, parameters);

                if (dt != null && dt.Rows.Count > 0)
                {
                    string empName = dt.Rows[0]["Name"].ToString();
                    string role = dt.Rows[0]["Role"].ToString();

                    MessageBox.Show($"أهلاً بك يا {empName} ({role})!", "تم تسجيل الدخول", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    sidebarPanel.Visible = true;
                    ShowPanel(dashboardPanel);
                    LoadDashboardData();
                }
                else
                {
                    MessageBox.Show("اسم المستخدم أو كلمة المرور غير صحيحة!", "خطأ في الدخول", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };

            card.Controls.AddRange(new Control[] { title, lblUser, txtUser, lblPass, txtPass, btnLogin });
            loginPanel.Controls.Add(card);
            mainContentPanel.Controls.Add(loginPanel);
        }

        private void ShowLoginScreen()
        {
            sidebarPanel.Visible = false;
            ShowPanel(loginPanel);
        }
        #endregion

        #region 2. Dashboard Screen (Live Data)
        private void BuildDashboardPanel()
        {
            dashboardPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(20) };

            Label lblTitle = new Label
            {
                Text = "Dashboard (Live Data)",
                Font = new Font("Segoe UI", 16, FontStyle.Bold),
                Location = new Point(20, 20),
                AutoSize = true
            };

            Panel card1 = CreateStatCard("Registered Patients", "0", Color.FromArgb(219, 234, 254), new Point(20, 60), out lblPatientCountVal);
            Panel card2 = CreateStatCard("Today's Visits", "0", Color.FromArgb(220, 252, 231), new Point(200, 60), out lblTodayVisitsVal);
            Panel card3 = CreateStatCard("Doctors", "0", Color.FromArgb(243, 232, 255), new Point(380, 60), out lblDoctorCountVal);
            Panel card4 = CreateStatCard("Nurses", "0", Color.FromArgb(254, 243, 199), new Point(560, 60), out lblNurseCountVal);

            Label lblTable = new Label
            {
                Text = "Recent Visits Record",
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                Location = new Point(20, 160),
                AutoSize = true
            };

            Button btnRefresh = new Button 
            { 
                Text = "🔄 Refresh Live Data", 
                Location = new Point(650, 155), 
                Width = 170, 
                Height = 30,
                BackColor = Color.FromArgb(30, 41, 59),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };

            dgvDashboardVisits = new DataGridView
            {
                Location = new Point(20, 190),
                Size = new Size(800, 420),
                BackgroundColor = Color.White,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                ReadOnly = true
            };

            btnRefresh.Click += (s, e) => LoadDashboardData();

            dashboardPanel.Controls.AddRange(new Control[] { lblTitle, card1, card2, card3, card4, lblTable, btnRefresh, dgvDashboardVisits });
            mainContentPanel.Controls.Add(dashboardPanel);
        }

        private void LoadDashboardData()
        {
            try
            {
                DataTable dtPatients = DatabaseHelper.ExecuteQuery("SELECT COUNT(*) FROM Patients");
                if (dtPatients != null && dtPatients.Rows.Count > 0)
                    lblPatientCountVal.Text = dtPatients.Rows[0][0].ToString();

                DataTable dtVisitsToday = DatabaseHelper.ExecuteQuery("SELECT COUNT(*) FROM Visits WHERE CAST(VisitDate AS DATE) = CAST(GETDATE() AS DATE)");
                if (dtVisitsToday != null && dtVisitsToday.Rows.Count > 0)
                    lblTodayVisitsVal.Text = dtVisitsToday.Rows[0][0].ToString();

                DataTable dtDoctors = DatabaseHelper.ExecuteQuery("SELECT COUNT(*) FROM Doctors");
                if (dtDoctors != null && dtDoctors.Rows.Count > 0)
                    lblDoctorCountVal.Text = dtDoctors.Rows[0][0].ToString();

                DataTable dtNurses = DatabaseHelper.ExecuteQuery("SELECT COUNT(*) FROM Nurses");
                if (dtNurses != null && dtNurses.Rows.Count > 0)
                    lblNurseCountVal.Text = dtNurses.Rows[0][0].ToString();

                string queryVisits = @"SELECT 
                                          V.VisitID AS [Visit ID], 
                                          P.Name AS [Patient Name], 
                                          S.Name AS [Doctor], 
                                          C.Name AS [Clinic], 
                                          V.VisitDate AS [Date & Time]
                                       FROM Visits V
                                       LEFT JOIN Patients P ON V.PatientID = P.PatientID
                                       LEFT JOIN Staff S ON V.DoctorID = S.EmployeeID
                                       LEFT JOIN Clinics C ON V.ClinicCode = C.ClinicCode
                                       ORDER BY V.VisitDate DESC";

                DataTable dtRecentVisits = DatabaseHelper.ExecuteQuery(queryVisits);
                dgvDashboardVisits.DataSource = dtRecentVisits;
            }
            catch (Exception ex)
            {
                MessageBox.Show("خطأ أثناء تحديث البيانات الحية: " + ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private Panel CreateStatCard(string title, string defaultValue, Color bgColor, Point location, out Label valueLabel)
        {
            Panel p = new Panel { Size = new Size(160, 80), BackColor = bgColor, Location = location };
            valueLabel = new Label { Text = defaultValue, Font = new Font("Segoe UI", 16, FontStyle.Bold), Location = new Point(10, 10), AutoSize = true };
            Label lblTxt = new Label { Text = title, Font = new Font("Segoe UI", 9, FontStyle.Regular), Location = new Point(10, 45), AutoSize = true };
            p.Controls.Add(valueLabel);
            p.Controls.Add(lblTxt);
            return p;
        }
        #endregion

        #region 3. Add Patient Screen (Placeholders Enabled)
        private void BuildAddPatientPanel()
        {
            addPatientPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(20) };

            Label lblTitle = new Label { Text = "Add Patient", Font = new Font("Segoe UI", 16, FontStyle.Bold), Location = new Point(20, 20), AutoSize = true };

            Label lblID = new Label { Text = "Patient ID", Location = new Point(30, 70) };
            TextBox txtID = new TextBox { Location = new Point(30, 95), Width = 250, Text = "" };
            SetPlaceholder(txtID, "Ex: P001234 (Required)");

            Label lblName = new Label { Text = "Full Name", Location = new Point(30, 130) };
            TextBox txtName = new TextBox { Location = new Point(30, 155), Width = 250, Text = "" };
            SetPlaceholder(txtName, "Ex: Fatima Ali (Required)");

            Label lblDOB = new Label { Text = "Date of Birth", Location = new Point(30, 190) };
            DateTimePicker dtpDOB = new DateTimePicker { Location = new Point(30, 215), Width = 250 };

            Label lblPhones = new Label { Text = "Phone Number", Location = new Point(320, 70) };
            TextBox txtPhone1 = new TextBox { Location = new Point(320, 95), Width = 250, Text = "" };
            SetPlaceholder(txtPhone1, "Ex: 0712345678 (Optional)");

            Button btnSave = new Button
            {
                Text = "Save Patient",
                Location = new Point(30, 280),
                Width = 120, Height = 35,
                BackColor = Color.FromArgb(37, 99, 235), ForeColor = Color.White, FlatStyle = FlatStyle.Flat
            };

            btnSave.Click += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(txtID.Text) || string.IsNullOrWhiteSpace(txtName.Text))
                {
                    MessageBox.Show("يرجى إدخال رقم المريض واسمه!", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string queryPatient = "INSERT INTO Patients (PatientID, Name, DateOfBirth) VALUES (@id, @name, @dob)";
                SqlParameter[] pPatient = {
                    new SqlParameter("@id", txtID.Text.Trim()),
                    new SqlParameter("@name", txtName.Text.Trim()),
                    new SqlParameter("@dob", dtpDOB.Value)
                };

                if (DatabaseHelper.ExecuteNonQuery(queryPatient, pPatient))
                {
                    if (!string.IsNullOrWhiteSpace(txtPhone1.Text))
                    {
                        string queryPhone = "INSERT INTO PatientPhones (PatientID, PhoneNumber) VALUES (@id, @phone)";
                        SqlParameter[] pPhone = {
                            new SqlParameter("@id", txtID.Text.Trim()),
                            new SqlParameter("@phone", txtPhone1.Text.Trim())
                        };
                        DatabaseHelper.ExecuteNonQuery(queryPhone, pPhone);
                    }

                    MessageBox.Show("تم حفظ بيانات المريض ورقم الهاتف بنجاح!", "نجاح", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txtID.Clear();
                    txtName.Clear();
                    txtPhone1.Clear();
                }
            };

            addPatientPanel.Controls.AddRange(new Control[] { lblTitle, lblID, txtID, lblName, txtName, lblDOB, dtpDOB, lblPhones, txtPhone1, btnSave });
            mainContentPanel.Controls.Add(addPatientPanel);
        }
        #endregion

        #region 4. Record Visit Screen (With Direct Diagnosis Linking)
        private void BuildRecordVisitPanel()
        {
            recordVisitPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(20) };

            Label lblTitle = new Label { Text = "Record Visit & Diagnosis", Font = new Font("Segoe UI", 16, FontStyle.Bold), Location = new Point(20, 20), AutoSize = true };

            // 1. رقم المريض
            Label lblPatient = new Label { Text = "Patient ID", Location = new Point(30, 70), AutoSize = true };
            TextBox txtPatientID = new TextBox { Location = new Point(30, 95), Width = 250, Text = "" };
            SetPlaceholder(txtPatientID, "Ex: P001234");

            // 2. رقم الطبيب
            Label lblDoctor = new Label { Text = "Doctor ID (Employee ID)", Location = new Point(30, 130), AutoSize = true };
            TextBox txtDoctorID = new TextBox { Location = new Point(30, 155), Width = 250, Text = "" };
            SetPlaceholder(txtDoctorID, "Ex: 1 (for Dr. Ahmed)");

            // 3. رمز العيادة
            Label lblClinic = new Label { Text = "Clinic Code", Location = new Point(300, 70), AutoSize = true };
            TextBox txtClinicCode = new TextBox { Location = new Point(300, 95), Width = 250, Text = "" };
            SetPlaceholder(txtClinicCode, "Ex: CLN01");

            // 4. رمز التشخيص (لربطه بالزيارة مباشرة)
            Label lblDiagnosis = new Label { Text = "Diagnosis Code", Location = new Point(300, 130), AutoSize = true };
            TextBox txtDiagnosisCode = new TextBox { Location = new Point(300, 155), Width = 250, Text = "" };
            SetPlaceholder(txtDiagnosisCode, "Ex: D101");

            // 5. تاريخ الزيارة
            Label lblDate = new Label { Text = "Visit Date", Location = new Point(30, 190), AutoSize = true };
            DateTimePicker dtpVisit = new DateTimePicker { Location = new Point(30, 215), Width = 250 };

            Button btnSaveVisit = new Button
            {
                Text = "Save Visit & Link Diagnosis",
                Location = new Point(300, 210),
                Width = 250,
                Height = 35,
                BackColor = Color.FromArgb(16, 185, 129),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };

            // حفظ الزيارات والربط المباشر بجدول Visit_Diagnoses
            btnSaveVisit.Click += (s, e) =>
            {
                string pId = txtPatientID.Text.Trim();
                string dId = txtDoctorID.Text.Trim();
                string cCode = txtClinicCode.Text.Trim();
                string diagCode = txtDiagnosisCode.Text.Trim();

                if (string.IsNullOrWhiteSpace(pId) || string.IsNullOrWhiteSpace(dId) || string.IsNullOrWhiteSpace(diagCode))
                {
                    MessageBox.Show("يرجى إدخال رقم المريض، رقم الطبيب، ورمز التشخيص!", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // استعلام لإدراج الزيارة ثم ربط رقم الزيارة المتولد تلقائياً برقم التشخيص
                string sqlQuery = @"
            INSERT INTO Visits (PatientID, DoctorID, ClinicCode, VisitDate) 
            VALUES (@pId, @dId, @cCode, @date);

            DECLARE @NewVisitID INT = SCOPE_IDENTITY();

            IF EXISTS (SELECT 1 FROM Diagnoses WHERE DiagnosisCode = @diagCode)
            BEGIN
                INSERT INTO Visit_Diagnoses (VisitID, DiagnosisCode) 
                VALUES (@NewVisitID, @diagCode);
            END";

                SqlParameter[] p = {
            new SqlParameter("@pId", pId),
            new SqlParameter("@dId", dId),
            new SqlParameter("@cCode", string.IsNullOrWhiteSpace(cCode) ? "CLN01" : cCode),
            new SqlParameter("@diagCode", diagCode),
            new SqlParameter("@date", dtpVisit.Value)
        };

                if (DatabaseHelper.ExecuteNonQuery(sqlQuery, p))
                {
                    MessageBox.Show($"تم تسجيل الزيارة وربطها بالتشخيص ({diagCode}) بنجاح!", "نجاح", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    // تفريغ الحقول
                    txtPatientID.Clear();
                    txtDoctorID.Clear();
                    txtClinicCode.Clear();
                    txtDiagnosisCode.Clear();
                }
                else
                {
                    MessageBox.Show("حدث خطأ أثناء الحفظ! التأكد من صحة رقم المريض ورمز التشخيص.", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };

            recordVisitPanel.Controls.AddRange(new Control[] {
        lblTitle, lblPatient, txtPatientID, lblDoctor, txtDoctorID,
        lblClinic, txtClinicCode, lblDiagnosis, txtDiagnosisCode,
        lblDate, dtpVisit, btnSaveVisit
    });

            mainContentPanel.Controls.Add(recordVisitPanel);
        }
        #endregion

        #region 5. Diagnoses Screen (Placeholders Enabled)
        private void BuildDiagnosesPanel()
        {
            diagnosesPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(20) };

            Label lblTitle = new Label { Text = "Diagnoses Management", Font = new Font("Segoe UI", 16, FontStyle.Bold), Location = new Point(20, 20), AutoSize = true };

            Label lblCode = new Label { Text = "Diagnosis Code", Location = new Point(30, 70) };
            TextBox txtCode = new TextBox { Location = new Point(30, 95), Width = 150, Text = "" };
            SetPlaceholder(txtCode, "Ex: D101");

            Label lblDesc = new Label { Text = "Description", Location = new Point(200, 70) };
            TextBox txtDesc = new TextBox { Location = new Point(200, 95), Width = 300, Text = "" };
            SetPlaceholder(txtDesc, "Ex: Seasonal Influenza");

            Button btnAdd = new Button { Text = "Add Diagnosis", Location = new Point(520, 93), Width = 120, Height = 27, BackColor = Color.FromArgb(37, 99, 235), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };

            DataGridView dgv = new DataGridView { Location = new Point(30, 150), Size = new Size(750, 420), BackgroundColor = Color.White, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, ReadOnly = true };

            btnAdd.Click += (s, e) =>
            {
                if (!string.IsNullOrWhiteSpace(txtCode.Text) && !string.IsNullOrWhiteSpace(txtDesc.Text))
                {
                    string query = "INSERT INTO Diagnoses (DiagnosisCode, Description) VALUES (@code, @desc)";
                    SqlParameter[] p = {
                        new SqlParameter("@code", txtCode.Text.Trim()),
                        new SqlParameter("@desc", txtDesc.Text.Trim())
                    };

                    if (DatabaseHelper.ExecuteNonQuery(query, p))
                    {
                        txtCode.Clear(); txtDesc.Clear();
                        MessageBox.Show("تم إدراج التشخيص في قاعدة البيانات بنجاح!", "نجاح", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        dgv.DataSource = DatabaseHelper.ExecuteQuery("SELECT DiagnosisCode AS [Code], Description FROM Diagnoses");
                    }
                }
                else
                {
                    MessageBox.Show("يرجى كتابة رمز التشخيص والوصف!", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            };

            diagnosesPanel.Controls.AddRange(new Control[] { lblTitle, lblCode, txtCode, lblDesc, txtDesc, btnAdd, dgv });
            mainContentPanel.Controls.Add(diagnosesPanel);
        }
        #endregion

        #region 6. Referrals Screen
        private void BuildReferralsPanel()
        {
            referralsPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(20) };

            Label lblTitle = new Label { Text = "Patient Referrals", Font = new Font("Segoe UI", 16, FontStyle.Bold), Location = new Point(20, 20), AutoSize = true };

            DataGridView dgv = new DataGridView { Location = new Point(30, 70), Size = new Size(750, 480), BackgroundColor = Color.White, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, ReadOnly = true };

            Button btnLoad = new Button { Text = "Load Referrals From SQL", Location = new Point(600, 20), Width = 180, Height = 35 };
            btnLoad.Click += (s, e) =>
            {
                string query = "SELECT ReferralID, PatientID, FromDoctorID, ToDoctorID, ReferralDate, Status FROM Referrals";
                dgv.DataSource = DatabaseHelper.ExecuteQuery(query);
            };

            referralsPanel.Controls.AddRange(new Control[] { lblTitle, btnLoad, dgv });
            mainContentPanel.Controls.Add(referralsPanel);
        }
        #endregion

        #region 7. Staff Screen
        private void BuildStaffPanel()
        {
            staffPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(20) };

            Label lblTitle = new Label { Text = "Staff & Doctors Management", Font = new Font("Segoe UI", 16, FontStyle.Bold), Location = new Point(20, 20), AutoSize = true };

            DataGridView dgv = new DataGridView { Location = new Point(30, 70), Size = new Size(780, 480), BackgroundColor = Color.White, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, ReadOnly = true };

            Button btnLoad = new Button { Text = "Load Staff Data", Location = new Point(630, 20), Width = 180, Height = 35 };
            btnLoad.Click += (s, e) =>
            {
                string query = "SELECT EmployeeID, Name, Role, YearsOfExperience, ClinicCode FROM Staff";
                dgv.DataSource = DatabaseHelper.ExecuteQuery(query);
            };

            staffPanel.Controls.AddRange(new Control[] { lblTitle, btnLoad, dgv });
            mainContentPanel.Controls.Add(staffPanel);
        }
        #endregion

        #region 8. Clinics Screen
        private void BuildClinicsPanel()
        {
            clinicsPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(20) };

            Label lblTitle = new Label { Text = "CityCare Clinics List", Font = new Font("Segoe UI", 16, FontStyle.Bold), Location = new Point(20, 20), AutoSize = true };

            DataGridView dgv = new DataGridView { Location = new Point(30, 70), Size = new Size(780, 480), BackgroundColor = Color.White, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, ReadOnly = true };

            Button btnLoad = new Button { Text = "Load Clinics Data", Location = new Point(630, 20), Width = 180, Height = 35 };
            btnLoad.Click += (s, e) =>
            {
                string query = "SELECT ClinicCode, Name, Address, Phone FROM Clinics";
                dgv.DataSource = DatabaseHelper.ExecuteQuery(query);
            };

            clinicsPanel.Controls.AddRange(new Control[] { lblTitle, btnLoad, dgv });
            mainContentPanel.Controls.Add(clinicsPanel);
        }
        #endregion

        #region 9. Reports Screen (Printing Adjusted - No Clipping)
        private void BuildReportsPanel()
        {
            reportsPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(20) };

            Label lblTitle = new Label { Text = "Patient Medical Report & Printing", Font = new Font("Segoe UI", 16, FontStyle.Bold), Location = new Point(20, 20), AutoSize = true };

            Label lblPatientID = new Label { Text = "Enter Patient ID:", Location = new Point(30, 70), AutoSize = true };
            TextBox txtPatientID = new TextBox { Location = new Point(140, 67), Width = 150, Text = "" };
            SetPlaceholder(txtPatientID, "Ex: P001234");

            Button btnGenerate = new Button 
            { 
                Text = "Generate Report", 
                Location = new Point(300, 65), 
                Width = 130, 
                Height = 28, 
                BackColor = Color.FromArgb(30, 41, 59), 
                ForeColor = Color.White, 
                FlatStyle = FlatStyle.Flat 
            };

            Button btnPrint = new Button 
            { 
                Text = "🖨️ Print / Save PDF", 
                Location = new Point(440, 65), 
                Width = 150, 
                Height = 28, 
                BackColor = Color.FromArgb(16, 185, 129), 
                ForeColor = Color.White, 
                FlatStyle = FlatStyle.Flat 
            };

            TextBox txtReportOutput = new TextBox
            {
                Location = new Point(30, 110),
                Size = new Size(780, 480),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("Consolas", 9),
                ReadOnly = true,
                BackColor = Color.White
            };

            // 1. توليد التقرير مع ربط التشخيصات بالزيارات حسب التاريخ
            btnGenerate.Click += (s, e) =>
            {
                string pID = txtPatientID.Text.Trim();
                if (string.IsNullOrEmpty(pID))
                {
                    MessageBox.Show("يرجى إدخال رقم المريض!", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // جلب بيانات المريض
                DataTable dtPatient = DatabaseHelper.ExecuteQuery("SELECT * FROM Patients WHERE PatientID = @id", new SqlParameter[] { new SqlParameter("@id", pID) });

                if (dtPatient == null || dtPatient.Rows.Count == 0)
                {
                    MessageBox.Show("لم يتم العثور على مريض بهذا الرقم!", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                string patientName = dtPatient.Rows[0]["Name"].ToString();
                string dob = Convert.ToDateTime(dtPatient.Rows[0]["DateOfBirth"]).ToString("yyyy-MM-dd");

                DataTable dtPhone = DatabaseHelper.ExecuteQuery("SELECT PhoneNumber FROM PatientPhones WHERE PatientID = @id", new SqlParameter[] { new SqlParameter("@id", pID) });
                string phone = (dtPhone.Rows.Count > 0) ? dtPhone.Rows[0]["PhoneNumber"].ToString() : "N/A";

                // استعلام يجلب الزيارات مع الطبيب والعيادة والتشخيص المربوط بها مرتبة حسب تاريخ الزيارة أحدثاً بأحدث
                string visitsWithDiagnosesQuery = @"
        SELECT 
            V.VisitID, 
            V.VisitDate, 
            S.Name AS Doctor, 
            C.Name AS Clinic,
            ISNULL(D.Description, 'No Diagnosis Recorded') AS Diagnosis
        FROM Visits V
        LEFT JOIN Staff S ON V.DoctorID = S.EmployeeID
        LEFT JOIN Clinics C ON V.ClinicCode = C.ClinicCode
        LEFT JOIN Visit_Diagnoses VD ON V.VisitID = VD.VisitID
        LEFT JOIN Diagnoses D ON VD.DiagnosisCode = D.DiagnosisCode
        WHERE V.PatientID = @id 
        ORDER BY V.VisitDate DESC";

                DataTable dtVisits = DatabaseHelper.ExecuteQuery(visitsWithDiagnosesQuery, new SqlParameter[] { new SqlParameter("@id", pID) });

                // صياغة التقرير النصي
                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                sb.AppendLine("=========================================================");
                sb.AppendLine("                 CITYCARE CLINIC SYSTEM                  ");
                sb.AppendLine("                 PATIENT MEDICAL REPORT                  ");
                sb.AppendLine("=========================================================");
                sb.AppendLine($" Report Date : {DateTime.Now:yyyy-MM-dd HH:mm}");
                sb.AppendLine("---------------------------------------------------------");
                sb.AppendLine(" PATIENT DETAILS:");
                sb.AppendLine($"   - Patient ID   : {pID}");
                sb.AppendLine($"   - Full Name    : {patientName}");
                sb.AppendLine($"   - Date of Birth: {dob}");
                sb.AppendLine($"   - Phone Number : {phone}");
                sb.AppendLine("---------------------------------------------------------");
                sb.AppendLine(" MEDICAL VISITS & DIAGNOSES HISTORY:");
                sb.AppendLine("---------------------------------------------------------");

                if (dtVisits != null && dtVisits.Rows.Count > 0)
                {
                    foreach (DataRow row in dtVisits.Rows)
                    {
                        string vID = row["VisitID"].ToString();
                        string vDate = Convert.ToDateTime(row["VisitDate"]).ToString("yyyy-MM-dd HH:mm");
                        string doc = row["Doctor"].ToString();
                        string clinic = row["Clinic"].ToString();
                        string diagnosis = row["Diagnosis"].ToString();

                        // عرض بيانات كل زيارة مع تشخيصها بشكل واضح
                        sb.AppendLine($" Visit ID   : {vID}  | Date: {vDate}");
                        sb.AppendLine($" Doctor     : {doc}");
                        sb.AppendLine($" Clinic     : {clinic}");
                        sb.AppendLine($" Diagnosis  : {diagnosis}");
                        sb.AppendLine("---------------------------------------------------------");
                    }
                }
                else
                {
                    sb.AppendLine(" No visit records found for this patient.");
                    sb.AppendLine("---------------------------------------------------------");
                }

                sb.AppendLine("             End of Report - CityCare Clinic             ");

                reportTextToPrint = sb.ToString();
                txtReportOutput.Text = reportTextToPrint;
            };

            printDoc.PrintPage += (s, e) =>
            {
                Font printFont = new Font("Consolas", 8.5f, FontStyle.Regular);
                float x = e.MarginBounds.Left - 20;
                float y = e.MarginBounds.Top - 20;
                float width = e.MarginBounds.Width + 40;
                float height = e.MarginBounds.Height + 40;

                RectangleF printArea = new RectangleF(x, y, width, height);

                StringFormat format = new StringFormat
                {
                    FormatFlags = StringFormatFlags.NoWrap
                };

                e.Graphics.DrawString(reportTextToPrint, printFont, Brushes.Black, printArea, format);
                e.HasMorePages = false;
            };

            btnPrint.Click += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(reportTextToPrint))
                {
                    MessageBox.Show("يرجى استخراج التقرير أولاً قبل الطباعة!", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                printDoc.DefaultPageSettings.Margins = new Margins(40, 40, 40, 40);

                PrintPreviewDialog previewDlg = new PrintPreviewDialog();
                previewDlg.Document = printDoc;
                previewDlg.Width = 850;
                previewDlg.Height = 650;
                previewDlg.ShowDialog();
            };

            reportsPanel.Controls.AddRange(new Control[] { lblTitle, lblPatientID, txtPatientID, btnGenerate, btnPrint, txtReportOutput });
            mainContentPanel.Controls.Add(reportsPanel);
        }
        #endregion
    }
}