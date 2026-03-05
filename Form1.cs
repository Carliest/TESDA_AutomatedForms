using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Color = System.Drawing.Color;
using Font = System.Drawing.Font;

namespace TESDAFormsApp;

// Custom rounded button class
public class RoundedButton : Button
{
    [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
    public int BorderRadius { get; set; } = 10;
    
    public RoundedButton()
    {
        FlatStyle = FlatStyle.Flat;
        FlatAppearance.BorderSize = 0;
    }
    
    protected override void OnSizeChanged(EventArgs e)
    {
        base.OnSizeChanged(e);
        SetButtonRegion();
    }
    
    private void SetButtonRegion()
    {
        using (var path = GetRoundedRectanglePath(new Rectangle(0, 0, Width - 1, Height - 1), BorderRadius))
        {
            Region = new Region(path);
        }
    }
    
    protected override void OnPaint(PaintEventArgs e)
    {
        e.Graphics.Clear(BackColor);
        e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
        
        using (var path = GetRoundedRectanglePath(new Rectangle(0, 0, Width - 1, Height - 1), BorderRadius))
        {
            using (var brush = new SolidBrush(BackColor))
            {
                e.Graphics.FillPath(brush, path);
            }
        }
        
        var stringFormat = new StringFormat
        {
            Alignment = StringAlignment.Center,
            LineAlignment = StringAlignment.Center
        };
        
        using (var brush = new SolidBrush(ForeColor))
        {
            e.Graphics.DrawString(Text, Font, brush, ClientRectangle, stringFormat);
        }
    }
    
    private System.Drawing.Drawing2D.GraphicsPath GetRoundedRectanglePath(Rectangle rect, int radius)
    {
        var path = new System.Drawing.Drawing2D.GraphicsPath();
        int diameter = radius * 2;
        
        path.AddArc(rect.X, rect.Y, diameter, diameter, 180, 90);
        path.AddArc(rect.Right - diameter, rect.Y, diameter, diameter, 270, 90);
        path.AddArc(rect.Right - diameter, rect.Bottom - diameter, diameter, diameter, 0, 90);
        path.AddArc(rect.X, rect.Bottom - diameter, diameter, diameter, 90, 90);
        path.CloseFigure();
        
        return path;
    }
}

// Assessment Fee mapping for each qualification
public static class AssessmentFees
{
    public static readonly Dictionary<string, Dictionary<string, decimal>> Fees = new()
    {
        { "Barangay Health Services NC II", new() { { "Full", 935.00m }, { "Absent", 0m } } },
        { "Bartending NC II", new() { { "Full", 1399.00m }, { "Absent", 0m } } },
        { "Bookkeeping NC II", new() { { "Full", 841.00m }, { "Absent", 0m } } },
        { "Bread and Pastry Production NC II", new() { { "Full", 1720.00m }, { "COC1", 1688.00m }, { "COC2", 1688.00m }, { "COC3", 1688.00m }, { "Absent", 0m } } },
        { "Computer System Servicing NC II", new() { { "Full", 1049.00m }, { "COC1", 863.00m }, { "COC2", 965.00m }, { "COC3", 859.00m }, { "COC4", 873.00m }, { "Absent", 0m } } },
        { "Cookery NC II", new() { { "Full", 1907.00m }, { "COC1", 1907.00m }, { "COC2", 1907.00m }, { "COC3", 1907.00m }, { "Absent", 0m } } },
        { "Driving NC II", new() { { "Full (Vehicle by Center)", 1034.00m }, { "Full (Vehicle by Candidate)", 819.00m }, { "Absent", 0m } } },
        { "Electrical Installation and Maintenance NC II", new() { { "Full", 1849.00m }, { "Absent", 0m } } },
        { "Event Management Services NC III", new() { { "Full", 905.00m }, { "COC1", 869.00m }, { "COC2", 848.00m }, { "Absent", 0m } } },
        { "Food and Beverage Services NC II", new() { { "Full", 882.00m }, { "Absent", 0m } } },
        { "Housekeeping NC II", new() { { "Full", 1108.00m }, { "COC1", 923.00m }, { "COC2", 997.00m }, { "COC3", 929.00m }, { "COC4", 963.00m }, { "Absent", 0m } } },
        { "Motorcycle/Small Engine Servicing NC II", new() { { "Full", 1491.00m }, { "COC1", 1199.00m }, { "Absent", 0m } } }
    };

    public static List<string> GetFeeOptions(string qualification)
    {
        if (Fees.ContainsKey(qualification))
            return Fees[qualification].Keys.ToList();
        return new() { "Full", "Absent" };
    }

    public static decimal GetFeeAmount(string qualification, string feeType)
    {
        if (Fees.ContainsKey(qualification) && Fees[qualification].ContainsKey(feeType))
            return Fees[qualification][feeType];
        return 0m;
    }
}

public class Candidate
{
    public string Name { get; set; } = "";
    public string Reference { get; set; } = "";
    public string AssessmentFee { get; set; } = "";
    public string AssessorFee { get; set; } = "";
}

public class Signatory
{
    public string AttendanceSheetName { get; set; } = "EDMAN L. VALENCIANO";
    public string AttendanceSheetPosition { get; set; } = "AC Manager";
    public string BillingAssessorACName { get; set; } = "EDMAN L. VALENCIANO";
    public string BillingAssessorACPosition { get; set; } = "AC Manager";
    public string BillingAssessorName { get; set; } = "RAMON C. SOLTES, JR.";
    public string BillingAssessorPosition { get; set; } = "AC Processing Officer";
    public string BillingAssessmentACName { get; set; } = "EDMAN L. VALENCIANO";
    public string BillingAssessmentACPosition { get; set; } = "AC Manager";
    public string BillingAssessmentVSName { get; set; } = "ROSALYN T. PERIDA, PhD";
    public string BillingAssessmentVSPosition { get; set; } = "Vocational School Superintendent I";
}

public class BatchInfo
{
    public int CandidateCount { get; set; }
    public string Qualification { get; set; } = "";
    public string Assessor { get; set; } = "";
    public string RQMCode { get; set; } = "";
    public string Scholarship { get; set; } = "";
    public string TrainingDuration { get; set; } = "";
    public string AssessmentDate { get; set; } = "";
    public List<BatchCandidate> Candidates { get; set; } = new();
    public BatchSignatory Signatories { get; set; } = new();
}

public class BatchSignatory
{
    public string AttendanceSheetName { get; set; } = "EDMAN L. VALENCIANO";
    public string AttendanceSheetPosition { get; set; } = "AC Manager";
    public string BillingAssessorACName { get; set; } = "EDMAN L. VALENCIANO";
    public string BillingAssessorACPosition { get; set; } = "AC Manager";
    public string BillingAssessorName { get; set; } = "RAMON C. SOLTES, JR.";
    public string BillingAssessorPosition { get; set; } = "AC Processing Officer";
    public string BillingAssessmentACName { get; set; } = "EDMAN L. VALENCIANO";
    public string BillingAssessmentACPosition { get; set; } = "AC Manager";
    public string BillingAssessmentVSName { get; set; } = "ROSALYN T. PERIDA, PhD";
    public string BillingAssessmentVSPosition { get; set; } = "Vocational School Superintendent I";
}

public class BatchCandidate
{
    public string Name { get; set; } = "";
    public string Reference { get; set; } = "";
    public string AssessmentFee { get; set; } = "Full";
    public string AssessorFee { get; set; } = "";
}

// ====================================================================
// MAIN FORM - Sidebar Navigation
// ====================================================================
public partial class MainForm : Form
{
    private Panel? sidebarPanel;
    private Panel? contentPanel;
    private Button? btnNewBatch;
    private Button? btnLoadBatch;
    private Button? btnReports;
    private NewBatchView? newBatchView;
    private LoadBatchView? loadBatchView;
    private ReportsView? reportsView;
    
    private readonly Color sidebarColor = Color.FromArgb(30, 41, 59);
    private readonly Color activeButtonColor = Color.FromArgb(37, 99, 235);
    private readonly Color hoverButtonColor = Color.FromArgb(51, 65, 85);
    private readonly Color inactiveTextColor = Color.FromArgb(203, 213, 225);
    
    public MainForm()
    {
        InitializeComponent();
        SetupForm();
        CreateSidebar();
        CreateContentArea();
        ShowNewBatchView();
    }
    
    private void SetupForm()
    {
        this.Text = "TESDA Forms Generator";
        this.Size = new Size(1400, 900);
        this.StartPosition = FormStartPosition.CenterScreen;
        this.BackColor = Color.FromArgb(241, 245, 249);
        this.Font = new Font("Segoe UI", 9F);
        this.FormBorderStyle = FormBorderStyle.Sizable;
        this.MinimumSize = new Size(1200, 700);
        
        try
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string iconPath = Path.Combine(baseDir, "Resources", "tesda-logo.png");
            
            if (File.Exists(iconPath))
            {
                using (var stream = new FileStream(iconPath, FileMode.Open))
                {
                    var bitmap = new Bitmap(stream);
                    IntPtr hIcon = bitmap.GetHicon();
                    this.Icon = Icon.FromHandle(hIcon);
                }
            }
        }
        catch { }
    }
    
    private void CreateSidebar()
    {
        sidebarPanel = new Panel
        {
            Width = 250,
            Dock = DockStyle.Left,
            BackColor = sidebarColor,
            Padding = new Padding(0)
        };
        
        Panel headerPanel = new Panel
        {
            Height = 80,
            Dock = DockStyle.Top,
            BackColor = sidebarColor,
            Padding = new Padding(20, 20, 20, 10)
        };
        
        Label titleLabel = new Label
        {
            Text = "TESDA Forms",
            Font = new Font("Segoe UI Semibold", 12F, FontStyle.Bold),
            ForeColor = Color.White,
            AutoSize = true,
            Location = new Point(20, 20)
        };
        
        Label subtitleLabel = new Label
        {
            Text = "Generator System",
            Font = new Font("Segoe UI", 8.5F),
            ForeColor = Color.FromArgb(148, 163, 184),
            AutoSize = true,
            Location = new Point(20, 45)
        };
        
        headerPanel.Controls.Add(titleLabel);
        headerPanel.Controls.Add(subtitleLabel);
        
        Panel navPanel = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = sidebarColor,
            Padding = new Padding(15)
        };
        
        btnNewBatch = CreateNavButton("New Batch", 15);
        btnLoadBatch = CreateNavButton("Load Batch", 65);
        btnReports = CreateNavButton("Reports", 115);
        
        navPanel.Controls.Add(btnNewBatch);
        navPanel.Controls.Add(btnLoadBatch);
        navPanel.Controls.Add(btnReports);
        
        Panel footerPanel = new Panel
        {
            Height = 50,
            Dock = DockStyle.Bottom,
            BackColor = sidebarColor,
            Padding = new Padding(20)
        };
        
        Label versionLabel = new Label
        {
            Text = "v2.0.0",
            Font = new Font("Segoe UI", 7.5F),
            ForeColor = Color.FromArgb(148, 163, 184),
            AutoSize = true,
            Location = new Point(20, 15)
        };
        
        footerPanel.Controls.Add(versionLabel);
        
        sidebarPanel.Controls.Add(navPanel);
        sidebarPanel.Controls.Add(headerPanel);
        sidebarPanel.Controls.Add(footerPanel);
        
        this.Controls.Add(sidebarPanel);
    }
    
    private Button CreateNavButton(string text, int yPosition)
    {
        Button btn = new Button
        {
            Text = text,
            Size = new Size(220, 45),
            Location = new Point(15, yPosition),
            FlatStyle = FlatStyle.Flat,
            BackColor = sidebarColor,
            ForeColor = inactiveTextColor,
            Font = new Font("Segoe UI Semibold", 9.5F, FontStyle.Bold),
            TextAlign = ContentAlignment.MiddleLeft,
            Padding = new Padding(15, 0, 0, 0),
            Cursor = Cursors.Hand
        };
        
        btn.FlatAppearance.BorderSize = 0;
        
        btn.MouseEnter += (s, e) => 
        {
            if (btn.BackColor != activeButtonColor)
            {
                btn.BackColor = hoverButtonColor;
                btn.ForeColor = Color.White;
            }
        };
        
        btn.MouseLeave += (s, e) => 
        {
            if (btn.BackColor != activeButtonColor)
            {
                btn.BackColor = sidebarColor;
                btn.ForeColor = inactiveTextColor;
            }
        };
        
        if (text == "New Batch")
            btn.Click += (s, e) => ShowNewBatchView();
        else if (text == "Load Batch")
            btn.Click += (s, e) => ShowLoadBatchView();
        else if (text == "Reports")
            btn.Click += (s, e) => ShowReportsView();
        
        return btn;
    }
    
    private void CreateContentArea()
    {
        contentPanel = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = Color.FromArgb(241, 245, 249),
            AutoScroll = true
        };
        
        this.Controls.Add(contentPanel);
    }
    
    private void SetActiveButton(Button? activeBtn)
    {
        if (btnNewBatch != null)
        {
            btnNewBatch.BackColor = sidebarColor;
            btnNewBatch.ForeColor = inactiveTextColor;
        }
        if (btnLoadBatch != null)
        {
            btnLoadBatch.BackColor = sidebarColor;
            btnLoadBatch.ForeColor = inactiveTextColor;
        }
        if (btnReports != null)
        {
            btnReports.BackColor = sidebarColor;
            btnReports.ForeColor = inactiveTextColor;
        }
        
        if (activeBtn != null)
        {
            activeBtn.BackColor = activeButtonColor;
            activeBtn.ForeColor = Color.White;
        }
    }
    
    private void ShowNewBatchView()
    {
        SetActiveButton(btnNewBatch);
        if (contentPanel != null)
            contentPanel.Controls.Clear();
        
        if (newBatchView == null)
            newBatchView = new NewBatchView();
        
        newBatchView.Dock = DockStyle.Fill;
        if (contentPanel != null)
            contentPanel.Controls.Add(newBatchView);
    }
    
    public void ShowNewBatchViewWithData(BatchInfo batchInfo)
    {
        SetActiveButton(btnNewBatch);
        if (contentPanel != null)
            contentPanel.Controls.Clear();
        
        if (newBatchView == null)
            newBatchView = new NewBatchView();
        
        newBatchView.LoadBatchData(batchInfo);
        newBatchView.Dock = DockStyle.Fill;
        if (contentPanel != null)
            contentPanel.Controls.Add(newBatchView);
    }
    
    private void ShowLoadBatchView()
    {
        SetActiveButton(btnLoadBatch);
        if (contentPanel != null)
            contentPanel.Controls.Clear();
        
        if (loadBatchView == null)
            loadBatchView = new LoadBatchView();
        
        loadBatchView.Dock = DockStyle.Fill;
        if (contentPanel != null)
            contentPanel.Controls.Add(loadBatchView);
    }
    
    private void ShowReportsView()
    {
        SetActiveButton(btnReports);
        if (contentPanel != null)
            contentPanel.Controls.Clear();
        
        if (reportsView == null)
            reportsView = new ReportsView();
        
        reportsView.Dock = DockStyle.Fill;
        if (contentPanel != null)
            contentPanel.Controls.Add(reportsView);
    }
    
    private void InitializeComponent()
    {
        this.SuspendLayout();
        this.ResumeLayout(false);
    }
}

// ====================================================================
// NEW BATCH VIEW
// ====================================================================
public partial class NewBatchView : UserControl
{
    NumericUpDown numCandidates = new NumericUpDown();
    ComboBox cboQualification = new ComboBox();
    TextBox txtAssessor = new TextBox();
    TextBox txtRQMCode = new TextBox();
    TextBox txtScholarship = new TextBox();
    TextBox txtTrainingDuration = new TextBox();
    DateTimePicker dtAssessment = new DateTimePicker();
    Button btnGenerate = new Button();
    Button btnAddCandidates = new Button();
    Button btnCustomDates = new Button();
    Button btnEdit = new Button();
    Button btnCreateAnother = new Button();
    Button btnEditSignatories = new Button();
    Label lblStatus = new Label();
    List<Candidate> customCandidates = new List<Candidate>();
    List<DateTime> customDates = new List<DateTime>();
    Signatory signatories = new Signatory();
    bool documentsGenerated = false;
    string currentOutputFolder = "";

    private readonly Color cardBackColor = Color.White;
    private readonly Color headerBackColor = Color.FromArgb(248, 250, 252);
    
    public NewBatchView()
    {
        InitializeComponent();
        SetupView();
    }
    
    private void SetupView()
    {
        this.Dock = DockStyle.Fill;
        this.BackColor = Color.FromArgb(241, 245, 249);
        this.AutoScroll = true;
        
        Panel topBar = CreateTopBar();
        topBar.Dock = DockStyle.Top;
        this.Controls.Add(topBar);
        
        Panel mainContent = new Panel
        {
            Dock = DockStyle.Fill,
            AutoScroll = false,
            BackColor = Color.FromArgb(241, 245, 249),
            Padding = new Padding(100)
        };
        
        int yPos = 100;
        
        Panel candidatesCard = CreateCandidatesCard();
        candidatesCard.Location = new Point(300, yPos);
        mainContent.Controls.Add(candidatesCard);
        yPos += candidatesCard.Height + 10;
        
        Panel qualificationCard = CreateQualificationCard();
        qualificationCard.Location = new Point(300, yPos);
        mainContent.Controls.Add(qualificationCard);
        yPos += qualificationCard.Height + 10;
        
        Panel adminCard = CreateAdminCard();
        adminCard.Location = new Point(300, yPos);
        mainContent.Controls.Add(adminCard);
        yPos += adminCard.Height + 10;
        
        Panel trainingCard = CreateTrainingCard();
        trainingCard.Location = new Point(300, yPos);
        mainContent.Controls.Add(trainingCard);
        yPos += trainingCard.Height + 10;
        
        Panel actionsPanel = CreateActionsPanel();
        actionsPanel.Location = new Point(300, yPos);
        mainContent.Controls.Add(actionsPanel);
        yPos += actionsPanel.Height + 10;
        
        lblStatus.Location = new Point(300, yPos);
        lblStatus.Width = 750;
        lblStatus.Height = 50;
        lblStatus.ForeColor = Color.FromArgb(100, 116, 139);
        lblStatus.Font = new Font("Segoe UI", 10);
        lblStatus.AutoSize = false;
        mainContent.Controls.Add(lblStatus);
        
        this.Controls.Add(mainContent);
    }
    
    private Panel CreateTopBar()
    {
        Panel topBar = new Panel
        {
            Height = 80,
            Dock = DockStyle.Top,
            BackColor = Color.White,
            Padding = new Padding(30, 15, 30, 15)
        };
        
        Label title = new Label
        {
            Text = "Create New Batch",
            Font = new Font("Segoe UI Semibold", 16F, FontStyle.Bold),
            ForeColor = Color.FromArgb(30, 41, 59),
            AutoSize = true,
            Location = new Point(295, 15)
        };
        
        Label subtitle = new Label
        {
            Text = "Fill in the details to generate TESDA forms",
            Font = new Font("Segoe UI", 9F),
            ForeColor = Color.FromArgb(100, 116, 139),
            AutoSize = true,
            Location = new Point(300, 55)
        };
        
        topBar.Controls.Add(title);
        topBar.Controls.Add(subtitle);
        
        return topBar;
    }
    
    private Panel CreateCard(string title, int width, int height)
    {
        Panel card = new Panel
        {
            Width = width,
            Height = height,
            BackColor = cardBackColor,
            BorderStyle = BorderStyle.FixedSingle,
            Anchor = AnchorStyles.Top | AnchorStyles.Left
        };
        
        Panel header = new Panel
        {
            Height = 50,
            Dock = DockStyle.Top,
            BackColor = headerBackColor
        };
        
        Label headerLabel = new Label
        {
            Text = title,
            Font = new Font("Segoe UI Semibold", 10F, FontStyle.Bold),
            ForeColor = Color.FromArgb(30, 41, 59),
            Location = new Point(20, 15),
            AutoSize = true
        };
        
        header.Controls.Add(headerLabel);
        card.Controls.Add(header);
        
        return card;
    }
    
    private Panel CreateCandidatesCard()
    {
        Panel card = CreateCard("Candidates Information", 750, 130);
        
        Label lblNum = new Label
        {
            Text = "Number of Candidates (Max 25)",
            Location = new Point(20, 50),
            AutoSize = true,
            Font = new Font("Segoe UI", 9F)
        };
        
        numCandidates.Location = new Point(20, 80);
        numCandidates.Width = 300;
        numCandidates.Maximum = 25;
        numCandidates.Minimum = 0;
        numCandidates.Font = new Font("Segoe UI", 10F);
        numCandidates.ValueChanged += (s, e) =>
        {
            numCandidates.BackColor = Color.White;
            btnCustomDates.Visible = (int)numCandidates.Value > 10;
        };
        
        btnAddCandidates.Text = "Enter Candidate Details";
        btnAddCandidates.Location = new Point(340, 80);
        btnAddCandidates.Width = 390;
        btnAddCandidates.Height = 32;
        btnAddCandidates.BackColor = Color.FromArgb(37, 99, 235);
        btnAddCandidates.ForeColor = Color.White;
        btnAddCandidates.FlatStyle = FlatStyle.Flat;
        btnAddCandidates.Font = new Font("Segoe UI Semibold", 9.5F, FontStyle.Bold);
        btnAddCandidates.Cursor = Cursors.Hand;
        btnAddCandidates.FlatAppearance.BorderSize = 0;
        btnAddCandidates.Click += (s, e) => OpenCandidateWindow();
        
        card.Controls.Add(lblNum);
        card.Controls.Add(numCandidates);
        card.Controls.Add(btnAddCandidates);
        
        return card;
    }
    
    private Panel CreateQualificationCard()
    {
        Panel card = CreateCard("Qualification & Assessor", 750, 175);
        
        Label lblQual = new Label
        {
            Text = "Qualification *",
            Location = new Point(20, 50),
            AutoSize = true,
            Font = new Font("Segoe UI", 9F)
        };
        
        cboQualification.Location = new Point(20, 70);
        cboQualification.Width = 710;
        cboQualification.DropDownStyle = ComboBoxStyle.DropDownList;
        cboQualification.Font = new Font("Segoe UI", 10F);
        cboQualification.Items.AddRange(new object[]
        {
            "Barangay Health Services NC II",
            "Bartending NC II",
            "Bookkeeping NC II",
            "Bread and Pastry Production NC II",
            "Computer System Servicing NC II",
            "Cookery NC II",
            "Driving NC II",
            "Electrical Installation and Maintenance NC II",
            "Event Management Services NC III",
            "Food and Beverage Services NC II",
            "Housekeeping NC II",
            "Motorcycle/Small Engine Servicing NC II"
        });
        cboQualification.SelectedIndexChanged += (s, e) => cboQualification.BackColor = Color.White;
        
        Label lblAssessor = new Label
        {
            Text = "Assessor Name *",
            Location = new Point(20, 110),
            AutoSize = true,
            Font = new Font("Segoe UI", 9F)
        };
        
        txtAssessor.Location = new Point(20, 130);
        txtAssessor.Width = 710;
        txtAssessor.Font = new Font("Segoe UI", 10F);
        txtAssessor.TextChanged += (s, e) => txtAssessor.BackColor = Color.White;
        
        card.Controls.Add(lblQual);
        card.Controls.Add(cboQualification);
        card.Controls.Add(lblAssessor);
        card.Controls.Add(txtAssessor);
        
        return card;
    }
    
    private Panel CreateAdminCard()
    {
        Panel card = CreateCard("Administrative Details", 750, 140);
        
        Label lblRQM = new Label
        {
            Text = "RQM Code *",
            Location = new Point(20, 70),
            AutoSize = true,
            Font = new Font("Segoe UI", 9F)
        };
        
        txtRQMCode.Location = new Point(20, 90);
        txtRQMCode.Width = 400;
        txtRQMCode.Font = new Font("Segoe UI", 10F);
        txtRQMCode.TextChanged += (s, e) => txtRQMCode.BackColor = Color.White;
        
        Label lblScholarship = new Label
        {
            Text = "Type of Scholarship *",
            Location = new Point(450, 70),
            AutoSize = true,
            Font = new Font("Segoe UI", 9F)
        };
        
        txtScholarship.Location = new Point(450, 90);
        txtScholarship.Width = 290;
        txtScholarship.Font = new Font("Segoe UI", 10F);
        txtScholarship.TextChanged += (s, e) => txtScholarship.BackColor = Color.White;
        
        card.Controls.Add(lblRQM);
        card.Controls.Add(txtRQMCode);
        card.Controls.Add(lblScholarship);
        card.Controls.Add(txtScholarship);
        
        return card;
    }
    
    private Panel CreateTrainingCard()
    {
        Panel card = CreateCard("Training & Assessment", 750, 140);
        
        Label lblDuration = new Label
        {
            Text = "Training Duration *",
            Location = new Point(20, 70),
            AutoSize = true,
            Font = new Font("Segoe UI", 9F)
        };
        
        txtTrainingDuration.Location = new Point(20, 90);
        txtTrainingDuration.Width = 400;
        txtTrainingDuration.Font = new Font("Segoe UI", 10F);
        txtTrainingDuration.PlaceholderText = "e.g., 3 months";
        txtTrainingDuration.TextChanged += (s, e) => txtTrainingDuration.BackColor = Color.White;
        
        Label lblDate = new Label
        {
            Text = "Assessment Date *",
            Location = new Point(450, 70),
            AutoSize = true,
            Font = new Font("Segoe UI", 9F)
        };
        
        dtAssessment.Location = new Point(450, 90);
        dtAssessment.Width = 290;
        dtAssessment.Font = new Font("Segoe UI", 10F);
        
        btnCustomDates.Text = "Custom Dates";
        btnCustomDates.Location = new Point(750, 90);
        btnCustomDates.Width = 100;
        btnCustomDates.Height = 32;
        btnCustomDates.BackColor = Color.FromArgb(100, 116, 139);
        btnCustomDates.ForeColor = Color.White;
        btnCustomDates.FlatStyle = FlatStyle.Flat;
        btnCustomDates.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
        btnCustomDates.Cursor = Cursors.Hand;
        btnCustomDates.FlatAppearance.BorderSize = 0;
        btnCustomDates.Visible = false;
        btnCustomDates.Click += (s, e) => OpenCustomDatesWindow();
        
        card.Controls.Add(lblDuration);
        card.Controls.Add(txtTrainingDuration);
        card.Controls.Add(lblDate);
        card.Controls.Add(dtAssessment);
        card.Controls.Add(btnCustomDates);
        
        return card;
    }
    
    private Panel CreateActionsPanel()
    {
        Panel panel = new Panel
        {
            Width = 750,
            Height = 140
        };
        
        btnGenerate.Text = "Create Batch";
        btnGenerate.Location = new Point(0, 0);
        btnGenerate.Width = 750;
        btnGenerate.Height = 45;
        btnGenerate.BackColor = Color.FromArgb(22, 163, 74);
        btnGenerate.ForeColor = Color.White;
        btnGenerate.FlatStyle = FlatStyle.Flat;
        btnGenerate.Font = new Font("Segoe UI Semibold", 10F, FontStyle.Bold);
        btnGenerate.Cursor = Cursors.Hand;
        btnGenerate.FlatAppearance.BorderSize = 0;
        btnGenerate.Click += Generate;
        
        btnEditSignatories.Text = "Edit Signatories";
        btnEditSignatories.Location = new Point(0, 55);
        btnEditSignatories.Width = 750;
        btnEditSignatories.Height = 40;
        btnEditSignatories.BackColor = Color.FromArgb(37, 99, 235);
        btnEditSignatories.ForeColor = Color.White;
        btnEditSignatories.FlatStyle = FlatStyle.Flat;
        btnEditSignatories.Font = new Font("Segoe UI Semibold", 10F, FontStyle.Bold);
        btnEditSignatories.Cursor = Cursors.Hand;
        btnEditSignatories.FlatAppearance.BorderSize = 0;
        btnEditSignatories.Click += (s, e) => OpenSignatoryEditor();
        
        btnCreateAnother.Text = "Create Another";
        btnCreateAnother.Location = new Point(0, 105);
        btnCreateAnother.Width = 440;
        btnCreateAnother.Height = 40;
        btnCreateAnother.BackColor = Color.FromArgb(103, 58, 183);
        btnCreateAnother.ForeColor = Color.White;
        btnCreateAnother.FlatStyle = FlatStyle.Flat;
        btnCreateAnother.Font = new Font("Segoe UI Semibold", 10F, FontStyle.Bold);
        btnCreateAnother.Cursor = Cursors.Hand;
        btnCreateAnother.FlatAppearance.BorderSize = 0;
        btnCreateAnother.Visible = false;
        btnCreateAnother.Click += (s, e) => ResetForm();
        
        btnEdit.Text = "Edit & Regenerate";
        btnEdit.Location = new Point(460, 105);
        btnEdit.Width = 440;
        btnEdit.Height = 40;
        btnEdit.BackColor = Color.FromArgb(255, 152, 0);
        btnEdit.ForeColor = Color.White;
        btnEdit.FlatStyle = FlatStyle.Flat;
        btnEdit.Font = new Font("Segoe UI Semibold", 10F, FontStyle.Bold);
        btnEdit.Cursor = Cursors.Hand;
        btnEdit.FlatAppearance.BorderSize = 0;
        btnEdit.Visible = false;
        btnEdit.Click += (s, e) => OpenEditForm();
        
        panel.Controls.Add(btnGenerate);
        panel.Controls.Add(btnEditSignatories);
        panel.Controls.Add(btnCreateAnother);
        panel.Controls.Add(btnEdit);
        
        return panel;
    }
    
    void OpenCandidateWindow()
    {
        int count = (int)numCandidates.Value;
        if (count == 0)
        {
            MessageBox.Show("Please enter number of candidates first", "Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        string qualification = cboQualification.SelectedItem?.ToString() ?? "";
        if (string.IsNullOrEmpty(qualification))
        {
            MessageBox.Show("Please select a qualification first", "Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        var candForm = new CandidateForm(count, customCandidates, qualification);
        if (candForm.ShowDialog() == DialogResult.OK)
        {
            customCandidates = candForm.GetCandidates();
        }
    }

    void OpenCustomDatesWindow()
    {
        int count = (int)numCandidates.Value;
        int dateCount = (int)Math.Ceiling(count / 10.0) - 1;
        var datesForm = new CustomDatesForm(dateCount, dtAssessment.Value, customDates);
        if (datesForm.ShowDialog() == DialogResult.OK)
        {
            customDates = datesForm.GetDates();
        }
    }

    void OpenEditForm()
    {
        var editForm = new EditForm(numCandidates, cboQualification, txtAssessor, txtRQMCode,
                                     txtScholarship, txtTrainingDuration, dtAssessment,
                                     customCandidates, customDates, signatories);
        if (editForm.ShowDialog() == DialogResult.OK)
        {
            customCandidates = editForm.GetCandidates();
            customDates = editForm.GetDates();
            signatories = editForm.GetSignatories();
            RegenerateDocuments();
        }
    }

    void OpenSignatoryEditor()
    {
        var sigForm = new SignatoryEditorForm(signatories);
        if (sigForm.ShowDialog() == DialogResult.OK)
        {
            signatories = sigForm.GetSignatories();
        }
    }

    void ResetForm()
    {
        numCandidates.Value = 0;
        cboQualification.SelectedIndex = -1;
        txtAssessor.Text = "";
        txtRQMCode.Text = "";
        txtScholarship.Text = "";
        txtTrainingDuration.Text = "";
        dtAssessment.Value = DateTime.Now;
        customCandidates.Clear();
        customDates.Clear();
        documentsGenerated = false;
        btnGenerate.Visible = true;
        btnCreateAnother.Visible = false;
        btnEdit.Visible = false;
        currentOutputFolder = "";
        lblStatus.Text = "";
        lblStatus.ForeColor = Color.Gray;
        btnCustomDates.Visible = false;
    }

    public void LoadBatchData(BatchInfo batchInfo)
    {
        try
        {
            // Populate form fields with batch data
            numCandidates.Value = batchInfo.CandidateCount;
            
            // Find and select the qualification
            for (int i = 0; i < cboQualification.Items.Count; i++)
            {
                if (cboQualification.Items[i].ToString() == batchInfo.Qualification)
                {
                    cboQualification.SelectedIndex = i;
                    break;
                }
            }
            
            txtAssessor.Text = batchInfo.Assessor;
            txtRQMCode.Text = batchInfo.RQMCode;
            txtScholarship.Text = batchInfo.Scholarship;
            txtTrainingDuration.Text = batchInfo.TrainingDuration;
            
            // Parse and set assessment date
            if (DateTime.TryParse(batchInfo.AssessmentDate, out DateTime assessmentDate))
            {
                dtAssessment.Value = assessmentDate;
            }
            
            // Load candidates
            customCandidates.Clear();
            foreach (var candidate in batchInfo.Candidates)
            {
                customCandidates.Add(new Candidate
                {
                    Name = candidate.Name,
                    Reference = candidate.Reference,
                    AssessmentFee = candidate.AssessmentFee,
                    AssessorFee = candidate.AssessorFee
                });
            }
            
            // Load signatories
            if (batchInfo.Signatories != null)
            {
                signatories = new Signatory
                {
                    AttendanceSheetName = batchInfo.Signatories.AttendanceSheetName,
                    AttendanceSheetPosition = batchInfo.Signatories.AttendanceSheetPosition,
                    BillingAssessorACName = batchInfo.Signatories.BillingAssessorACName,
                    BillingAssessorACPosition = batchInfo.Signatories.BillingAssessorACPosition,
                    BillingAssessorName = batchInfo.Signatories.BillingAssessorName,
                    BillingAssessorPosition = batchInfo.Signatories.BillingAssessorPosition,
                    BillingAssessmentACName = batchInfo.Signatories.BillingAssessmentACName,
                    BillingAssessmentACPosition = batchInfo.Signatories.BillingAssessmentACPosition,
                    BillingAssessmentVSName = batchInfo.Signatories.BillingAssessmentVSName,
                    BillingAssessmentVSPosition = batchInfo.Signatories.BillingAssessmentVSPosition
                };
            }
            
            lblStatus.Text = "✓ Batch loaded successfully";
            lblStatus.ForeColor = Color.Green;
        }
        catch (Exception ex)
        {
            lblStatus.Text = $"✗ Error loading batch: {ex.Message}";
            lblStatus.ForeColor = Color.Red;
            MessageBox.Show($"Error loading batch: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    bool ValidateAllFields()
    {
        bool isValid = true;
        List<string> errors = new List<string>();

        numCandidates.BackColor = Color.White;
        cboQualification.BackColor = Color.White;
        txtAssessor.BackColor = Color.White;
        txtRQMCode.BackColor = Color.White;
        txtScholarship.BackColor = Color.White;
        txtTrainingDuration.BackColor = Color.White;

        int N = (int)numCandidates.Value;
        if (N == 0)
        {
            numCandidates.BackColor = Color.FromArgb(255, 200, 200);
            errors.Add("Number of candidates is required");
            isValid = false;
        }

        if (cboQualification.SelectedIndex < 0)
        {
            cboQualification.BackColor = Color.FromArgb(255, 200, 200);
            errors.Add("Qualification must be selected");
            isValid = false;
        }

        if (string.IsNullOrWhiteSpace(txtAssessor.Text))
        {
            txtAssessor.BackColor = Color.FromArgb(255, 200, 200);
            errors.Add("Assessor Name is required");
            isValid = false;
        }

        if (string.IsNullOrWhiteSpace(txtRQMCode.Text))
        {
            txtRQMCode.BackColor = Color.FromArgb(255, 200, 200);
            errors.Add("RQM Code is required");
            isValid = false;
        }

        if (string.IsNullOrWhiteSpace(txtScholarship.Text))
        {
            txtScholarship.BackColor = Color.FromArgb(255, 200, 200);
            errors.Add("Type of Scholarship/Modality is required");
            isValid = false;
        }

        if (string.IsNullOrWhiteSpace(txtTrainingDuration.Text))
        {
            txtTrainingDuration.BackColor = Color.FromArgb(255, 200, 200);
            errors.Add("Training Duration is required");
            isValid = false;
        }

        if (N > 0)
        {
            if (customCandidates.Count == 0)
            {
                errors.Add("Candidate Details are required - Click 'Enter Candidates' to add candidates");
                isValid = false;
            }
            else if (customCandidates.Count < N)
            {
                errors.Add($"Only {customCandidates.Count} out of {N} candidate(s) entered - please fill in all candidate details");
                isValid = false;
            }
            else
            {
                bool hasMissingData = false;
                for (int i = 0; i < customCandidates.Count; i++)
                {
                    var candidate = customCandidates[i];
                    if (string.IsNullOrWhiteSpace(candidate.Name) || candidate.Name.Contains("Candidate"))
                    {
                        hasMissingData = true;
                        break;
                    }
                    if (string.IsNullOrWhiteSpace(candidate.Reference) || candidate.Reference.Contains("AC-"))
                    {
                        hasMissingData = true;
                        break;
                    }
                }

                if (hasMissingData)
                {
                    errors.Add("Some candidate names or references are still blank - please fill in all candidate details");
                    isValid = false;
                }
            }
        }

        if (!isValid)
        {
            lblStatus.ForeColor = Color.Red;
            lblStatus.Text = "⚠ Please fill in all required fields";
            string errorMessage = string.Join("\n", errors);
            MessageBox.Show(errorMessage, "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        return isValid;
    }

    void Generate(object? sender, EventArgs e)
    {
        if (!ValidateAllFields())
        {
            return;
        }

        int N = (int)numCandidates.Value;

        try
        {
            lblStatus.ForeColor = Color.Blue;
            lblStatus.Text = "Generating forms...";
            this.Refresh();

            List<Candidate> candidates = new List<Candidate>();
            if (customCandidates.Count >= N)
            {
                candidates = customCandidates.Take(N).ToList();
            }
            else if (customCandidates.Count > 0)
            {
                candidates = new List<Candidate>(customCandidates);
                for (int i = customCandidates.Count; i < N; i++)
                {
                    candidates.Add(new Candidate()
                    {
                        Name = "Candidate " + (i + 1),
                        Reference = "AC-" + (i + 1).ToString("000"),
                        AssessmentFee = ""
                    });
                }
            }
            else
            {
                for (int i = 0; i < N; i++)
                {
                    candidates.Add(new Candidate()
                    {
                        Name = "Candidate " + (i + 1),
                        Reference = "AC-" + (i + 1).ToString("000"),
                        AssessmentFee = "",
                        AssessorFee = ""
                    });
                }
            }

            using (var titleForm = new GenerationTitleForm())
            {
                if (titleForm.ShowDialog() != DialogResult.OK)
                {
                    lblStatus.ForeColor = Color.Gray;
                    lblStatus.Text = "Generation cancelled";
                    return;
                }

                string folderTitle = titleForm.FolderTitle;
                currentOutputFolder = Path.Combine("Output", folderTitle);

                if (Directory.Exists(currentOutputFolder))
                {
                    DialogResult result = MessageBox.Show(
                        $"A folder with the name '{folderTitle}' already exists.\n\nDo you want to overwrite the existing files?",
                        "Folder Already Exists",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning
                    );

                    if (result != DialogResult.Yes)
                    {
                        lblStatus.ForeColor = Color.Gray;
                        lblStatus.Text = "Generation cancelled";
                        return;
                    }
                }

                Directory.CreateDirectory(currentOutputFolder);
            }

            GenerateBilling("Templates/BILL-ASSESSOR'S FEE.docx", Path.Combine(currentOutputFolder, "Billing_AssessorFee.docx"), candidates, N);
            GenerateBilling("Templates/Billing Assessment.docx", Path.Combine(currentOutputFolder, "Billing_AssessmentFee.docx"), candidates, N);
            GenerateAttendance("Templates/Attendance Sheet.docx", Path.Combine(currentOutputFolder, "Attendance_Sheet.docx"), candidates, dtAssessment.Value, customDates);

            SaveBatchInfo(N);

            documentsGenerated = true;
            btnGenerate.Visible = false;
            btnCreateAnother.Visible = true;
            btnEdit.Visible = true;
            lblStatus.ForeColor = Color.Green;
            lblStatus.Text = "✓ Documents Generated Successfully! Click 'Edit & Regenerate' to make changes.";
            MessageBox.Show($"Documents Generated Successfully\n\nFiles saved in '{currentOutputFolder}' folder\n\nYou can now use 'Edit & Regenerate' button to make changes.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            lblStatus.ForeColor = Color.Red;
            lblStatus.Text = "✗ Error generating documents";
            MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    void RegenerateDocuments()
    {
        if (!ValidateAllFields())
        {
            return;
        }

        int N = (int)numCandidates.Value;

        try
        {
            lblStatus.ForeColor = Color.Blue;
            lblStatus.Text = "Regenerating documents...";
            this.Refresh();

            List<Candidate> candidates = new List<Candidate>();
            if (customCandidates.Count >= N)
            {
                candidates = customCandidates.Take(N).ToList();
            }
            else if (customCandidates.Count > 0)
            {
                candidates = new List<Candidate>(customCandidates);
                for (int i = customCandidates.Count; i < N; i++)
                {
                    candidates.Add(new Candidate()
                    {
                        Name = "Candidate " + (i + 1),
                        Reference = "AC-" + (i + 1).ToString("000"),
                        AssessmentFee = "",
                        AssessorFee = ""
                    });
                }
            }
            else
            {
                for (int i = 0; i < N; i++)
                {
                    candidates.Add(new Candidate()
                    {
                        Name = "Candidate " + (i + 1),
                        Reference = "AC-" + (i + 1).ToString("000"),
                        AssessmentFee = "",
                        AssessorFee = ""
                    });
                }
            }

            DeletePreviousFiles();

            GenerateBilling("Templates/BILL-ASSESSOR'S FEE.docx", Path.Combine(currentOutputFolder, "Billing_AssessorFee.docx"), candidates, N);
            GenerateBilling("Templates/Billing Assessment.docx", Path.Combine(currentOutputFolder, "Billing_AssessmentFee.docx"), candidates, N);
            GenerateAttendance("Templates/Attendance Sheet.docx", Path.Combine(currentOutputFolder, "Attendance_Sheet.docx"), candidates, dtAssessment.Value, customDates);

            lblStatus.ForeColor = Color.Green;
            lblStatus.Text = "✓ Documents Updated Successfully!";
            MessageBox.Show($"Documents Updated Successfully\n\nFiles saved in '{currentOutputFolder}' folder", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            lblStatus.ForeColor = Color.Red;
            lblStatus.Text = "✗ Error regenerating documents";
            MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    void DeletePreviousFiles()
    {
        try
        {
            if (string.IsNullOrEmpty(currentOutputFolder))
                return;

            string[] filesToDelete = {
                Path.Combine(currentOutputFolder, "Billing_AssessorFee.docx"),
                Path.Combine(currentOutputFolder, "Billing_AssessmentFee.docx"),
                Path.Combine(currentOutputFolder, "Attendance_Sheet.docx")
            };

            foreach (var file in filesToDelete)
            {
                if (File.Exists(file))
                {
                    File.Delete(file);
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Warning: Could not delete previous files: {ex.Message}");
        }
    }

    void SaveBatchInfo(int candidateCount)
    {
        try
        {
            var batchInfo = new BatchInfo
            {
                CandidateCount = candidateCount,
                Qualification = cboQualification.SelectedItem?.ToString() ?? "",
                Assessor = txtAssessor.Text,
                RQMCode = txtRQMCode.Text,
                Scholarship = txtScholarship.Text,
                TrainingDuration = txtTrainingDuration.Text,
                AssessmentDate = dtAssessment.Value.ToString("yyyy-MM-dd"),
                Candidates = customCandidates.Select(c => new BatchCandidate { Name = c.Name, Reference = c.Reference, AssessmentFee = c.AssessmentFee, AssessorFee = c.AssessorFee }).ToList(),
                Signatories = new BatchSignatory
                {
                    AttendanceSheetName = signatories.AttendanceSheetName,
                    AttendanceSheetPosition = signatories.AttendanceSheetPosition,
                    BillingAssessorACName = signatories.BillingAssessorACName,
                    BillingAssessorACPosition = signatories.BillingAssessorACPosition,
                    BillingAssessorName = signatories.BillingAssessorName,
                    BillingAssessorPosition = signatories.BillingAssessorPosition,
                    BillingAssessmentACName = signatories.BillingAssessmentACName,
                    BillingAssessmentACPosition = signatories.BillingAssessmentACPosition,
                    BillingAssessmentVSName = signatories.BillingAssessmentVSName,
                    BillingAssessmentVSPosition = signatories.BillingAssessmentVSPosition
                }
            };

            string batchInfoPath = Path.Combine(currentOutputFolder, "batch_info.json");
            string json = JsonSerializer.Serialize(batchInfo, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(batchInfoPath, json);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Warning: Could not save batch info: {ex.Message}");
        }
    }

    void GenerateBilling(string template, string output, List<Candidate> candidates, int count)
    {
        if (!File.Exists(template))
        {
            MessageBox.Show($"Template not found: {template}");
            return;
        }

        try
        {
            File.Copy(template, output, true);
            using (WordprocessingDocument doc = WordprocessingDocument.Open(output, true))
            {
                var tables = doc.MainDocumentPart?.Document?.Body?.Elements<Table>().ToList();
                if (tables == null || tables.Count == 0)
                {
                    MessageBox.Show($"No table found in template: {template}");
                    return;
                }

                Table table = tables.First();
                var rows = table.Elements<TableRow>().ToList();

                bool isAssessmentForm = template.Contains("Billing Assessment");
                int rowsToKeep = isAssessmentForm ? 27 : count + 2;

                while (rows.Count > rowsToKeep)
                {
                    table.RemoveChild(rows[rows.Count - 1]);
                    rows.RemoveAt(rows.Count - 1);
                }

                int r = 1;
                decimal totalFee = 0;
                string qualification = cboQualification.SelectedItem?.ToString() ?? "";
                
                foreach (var c in candidates)
                {
                    if (r < rows.Count)
                    {
                        var cells = rows[r].Elements<TableCell>().ToList();
                        if (cells.Count > 2)
                        {
                            cells[1].RemoveAllChildren<Paragraph>();
                            cells[1].Append(new Paragraph(new Run(new Text(c.Name))));
                            cells[2].RemoveAllChildren<Paragraph>();
                            cells[2].Append(new Paragraph(new Run(new Text(c.Reference))));
                            
                            if (cells.Count > 3)
                            {
                                string feeDisplay = "";
                                
                                if (isAssessmentForm)
                                {
                                    if (!string.IsNullOrEmpty(c.AssessmentFee) && c.AssessmentFee != "Absent")
                                    {
                                        decimal feeAmount = AssessmentFees.GetFeeAmount(qualification, c.AssessmentFee);
                                        feeDisplay = feeAmount.ToString("F2");
                                        totalFee += feeAmount;
                                    }
                                }
                                else
                                {
                                    if (!string.IsNullOrEmpty(c.AssessorFee))
                                    {
                                        if (decimal.TryParse(c.AssessorFee, out decimal assessorFeeAmount))
                                        {
                                            feeDisplay = assessorFeeAmount.ToString("F2");
                                            totalFee += assessorFeeAmount;
                                        }
                                    }
                                }
                                
                                cells[3].RemoveAllChildren<Paragraph>();
                                cells[3].Append(new Paragraph(new Run(new Text(feeDisplay))));
                            }
                        }
                        r++;
                    }
                }
                
                if (rows.Count > 0)
                {
                    var totalRow = rows[rows.Count - 1];
                    totalRow.RemoveAllChildren<TableCell>();
                    
                    var headerRow = rows[0];
                    var headerCells = headerRow.Elements<TableCell>().ToList();
                    Shading? headerShading = null;
                    if (headerCells.Count > 0)
                    {
                        headerShading = headerCells[0].TableCellProperties?.Elements<Shading>().FirstOrDefault();
                    }
                    
                    var cell1 = new TableCell();
                    var tcPr1 = new TableCellProperties();
                    tcPr1.Append(new GridSpan() { Val = 3 });
                    if (headerShading != null)
                    {
                        var shading1 = new Shading() { Fill = headerShading.Fill };
                        tcPr1.Append(shading1);
                    }
                    cell1.Append(tcPr1);
                    cell1.Append(new Paragraph(new Run(new Text("Total"))));
                    totalRow.Append(cell1);
                    
                    var cell2 = new TableCell();
                    var tcPr2 = new TableCellProperties();
                    tcPr2.Append(new GridSpan() { Val = 1 });
                    if (headerShading != null)
                    {
                        var shading2 = new Shading() { Fill = headerShading.Fill };
                        tcPr2.Append(shading2);
                    }
                    cell2.Append(tcPr2);
                    if (totalFee > 0)
                    {
                        cell2.Append(new Paragraph(new Run(new Text(totalFee.ToString("F2")))));
                    }
                    else
                    {
                        cell2.Append(new Paragraph());
                    }
                    totalRow.Append(cell2);
                }

                ReplaceFieldValue(doc, "1", cboQualification.SelectedItem?.ToString() ?? "");
                ReplaceFieldValue(doc, "2", txtAssessor.Text);
                ReplaceFieldValue(doc, "3", txtRQMCode.Text);
                ReplaceFieldValue(doc, "4", dtAssessment.Value.ToString("MM/dd/yyyy"));
                ReplaceFieldValue(doc, "5", txtScholarship.Text);
                ReplaceFieldValue(doc, "6", txtTrainingDuration.Text);

                if (template.Contains("ASSESSOR") || template.Contains("Assessor"))
                {
                    ReplaceSignatory(doc, "EDMAN L. VALENCIANO", signatories.BillingAssessorACName, true);
                    ReplaceSignatory(doc, "AC Manager", signatories.BillingAssessorACPosition, false);
                    ReplaceSignatory(doc, "RAMON C. SOLTES, JR", signatories.BillingAssessorName, true);
                    ReplaceSignatory(doc, "AC Processing Officer", signatories.BillingAssessorPosition, false);
                }
                else if (template.Contains("Assessment") || template.Contains("ASSESSMENT"))
                {
                    ReplaceSignatory(doc, "EDMAN L. VALENCIANO", signatories.BillingAssessmentACName, true);
                    ReplaceSignatory(doc, "AC Manager", signatories.BillingAssessmentACPosition, false);
                    ReplaceSignatory(doc, "ROSALYN T. PERIDA, PhD", signatories.BillingAssessmentVSName, true);
                    ReplaceSignatory(doc, "Vocational School Superintendent I", signatories.BillingAssessmentVSPosition, false);
                }

                doc.MainDocumentPart?.Document?.Save();
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error generating billing document: {ex.Message}");
        }
    }

    void GenerateAttendance(string template, string output, List<Candidate> candidates, DateTime startDate, List<DateTime> customDates = null!)
    {
        if (!File.Exists(template))
        {
            MessageBox.Show($"Template not found: {template}");
            return;
        }

        try
        {
            File.Copy(template, output, true);

            using (WordprocessingDocument doc = WordprocessingDocument.Open(output, true))
            {
                var body = doc.MainDocumentPart?.Document?.Body;
                var tables = body?.Elements<Table>().ToList();
                if (tables == null || tables.Count == 0)
                {
                    MessageBox.Show("No table found in attendance template");
                    return;
                }

                int numPages = candidates.Count <= 10 ? 1 : (candidates.Count <= 20 ? 2 : 3);
                
                var bodyElements = body?.ChildElements.Cast<OpenXmlElement>().ToList() ?? new List<OpenXmlElement>();
                
                var table = tables.Last();
                var rows = table.Elements<TableRow>().ToList();

                ReplaceFieldValue(doc, "QUALIFICATION TITLE", cboQualification.SelectedItem?.ToString() ?? "");
                ReplaceFieldValue(doc, "1", cboQualification.SelectedItem?.ToString() ?? "");
                ReplaceFieldValue(doc, "2", txtAssessor.Text);
                ReplaceFieldValue(doc, "3", txtRQMCode.Text);
                ReplaceFieldValue(doc, "5", txtScholarship.Text);
                ReplaceFieldValue(doc, "6", txtTrainingDuration.Text);
                
                ReplaceSignatory(doc, "EDMAN L. VALENCIANO", signatories.AttendanceSheetName, true);
                ReplaceSignatory(doc, "AC Manager", signatories.AttendanceSheetPosition, false);

                void PopulateTablePage(Table tbl, int startCandidateIndex, DateTime pageDate)
                {
                    var tblRows = tbl.Elements<TableRow>().ToList();
                    
                    UpdateDateInTable(tbl, pageDate);
                    
                    for (int i = 0; i < 10 && (i + 1) < tblRows.Count; i++)
                    {
                        int rowIdx = i + 3;
                        int candIdx = startCandidateIndex + i;
                        int rowNum = i + 1;
                        
                        var cells = tblRows[rowIdx].Elements<TableCell>().ToList();
                        if (cells.Count >= 3)
                        {
                            var noCell = cells[0];
                            var noParagraphs = noCell.Elements<Paragraph>().ToList();
                            if (noParagraphs.Count > 0)
                            {
                                noParagraphs[0].RemoveAllChildren();
                                noParagraphs[0].Append(new Run(new Text(rowNum.ToString())));
                            }
                            
                            var nameCell = cells[1];
                            var nameParagraphs = nameCell.Elements<Paragraph>().ToList();
                            if (nameParagraphs.Count > 0)
                            {
                                nameParagraphs[0].RemoveAllChildren();
                                if (candIdx < candidates.Count)
                                {
                                    nameParagraphs[0].Append(new Run(new Text(candidates[candIdx].Name)));
                                }
                            }
                            
                            var refCell = cells[2];
                            var refParagraphs = refCell.Elements<Paragraph>().ToList();
                            if (refParagraphs.Count > 0)
                            {
                                refParagraphs[0].RemoveAllChildren();
                                if (candIdx < candidates.Count)
                                {
                                    refParagraphs[0].Append(new Run(new Text(candidates[candIdx].Reference)));
                                }
                            }
                        }
                    }
                }

                void UpdateDateInTable(Table tbl, DateTime date)
                {
                    var tblRows = tbl.Elements<TableRow>().ToList();
                    for (int r = 0; r < Math.Min(3, tblRows.Count); r++)
                    {
                        var cells = tblRows[r].Elements<TableCell>().ToList();
                        for (int c = 0; c < cells.Count; c++)
                        {
                            var cellText = cells[c].InnerText;
                            if (cellText.Contains("Date of Assessment") && c + 1 < cells.Count)
                            {
                                var dateCell = cells[c + 1];
                                var paras = dateCell.Elements<Paragraph>().ToList();
                                if (paras.Count > 0)
                                {
                                    paras[0].RemoveAllChildren();
                                    paras[0].Append(new Run(new Text(date.ToString("MM/dd/yyyy"))));
                                }
                                return;
                            }
                        }
                    }
                }

                PopulateTablePage(table, 0, startDate);

                if (numPages > 1)
                {
                    for (int pageNum = 1; pageNum < numPages; pageNum++)
                    {
                        DateTime pageDate;
                        if (customDates != null && customDates.Count > 0 && pageNum - 1 < customDates.Count)
                        {
                            pageDate = customDates[pageNum - 1];
                        }
                        else
                        {
                            pageDate = startDate.AddDays(pageNum);
                        }

                        foreach (var element in bodyElements)
                        {
                            body?.Append((OpenXmlElement)element.CloneNode(true));
                        }
                        
                        var allTables = body?.Elements<Table>().ToList();
                        if (allTables != null && allTables.Count > 0)
                        {
                            var clonedTable = allTables.Last();
                            int startCandidate = pageNum * 10;
                            PopulateTablePage(clonedTable, startCandidate, pageDate);
                        }
                    }
                }

                doc.MainDocumentPart?.Document?.Save();
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error generating attendance sheet: {ex.Message}");
        }
    }

    void ReplaceFieldValue(WordprocessingDocument doc, string placeholder, string value)
    {
        try
        {
            var allTextElements = doc.MainDocumentPart?.Document?.Descendants<Text>().ToList();
            if (allTextElements == null) return;

            foreach (var textElem in allTextElements)
            {
                if (textElem.Text == placeholder)
                {
                    textElem.Text = value;

                    var run = textElem.Ancestors<Run>().FirstOrDefault();
                    if (run != null)
                    {
                        var runProps = run.Elements<RunProperties>().FirstOrDefault();
                        if (runProps == null)
                        {
                            runProps = new RunProperties();
                            run.InsertAt(runProps, 0);
                        }
                        runProps.RemoveAllChildren<Bold>();
                        runProps.Append(new Bold());
                    }
                    return;
                }
            }

            string fieldName = placeholder switch
            {
                "1" => "Qualification",
                "2" => "Assessor",
                "3" => "RQM Code",
                "4" => "Date of Assessment",
                "5" => "Type of Scholarship/Modality",
                "6" => "Training Duration",
                _ => placeholder
            };

            var allParas = doc.MainDocumentPart?.Document?.Body?.Descendants<Paragraph>().ToList();
            if (allParas == null) return;
            foreach (var para in allParas)
            {
                var fullText = para.InnerText;

                if (!fullText.Contains(fieldName + ":"))
                    continue;

                var runs = para.Elements<Run>().ToList();

                int colonRunIdx = -1;
                for (int i = 0; i < runs.Count; i++)
                {
                    if (runs[i].InnerText.Contains(":"))
                    {
                        string textBeforeColon = fullText.Substring(0, fullText.IndexOf(runs[i].InnerText) + runs[i].InnerText.IndexOf(":") + 1);
                        if (textBeforeColon.Contains(fieldName + ":"))
                        {
                            colonRunIdx = i;
                            break;
                        }
                    }
                }

                if (colonRunIdx == -1)
                    continue;

                string valueWithSpace = " " + value;

                if (colonRunIdx + 1 < runs.Count)
                {
                    runs[colonRunIdx + 1].RemoveAllChildren<Text>();
                    runs[colonRunIdx + 1].Append(new Text(valueWithSpace));

                    var runProps = runs[colonRunIdx + 1].Elements<RunProperties>().FirstOrDefault();
                    if (runProps == null)
                    {
                        runProps = new RunProperties();
                        runs[colonRunIdx + 1].InsertAt(runProps, 0);
                    }
                    runProps.RemoveAllChildren<Bold>();
                    runProps.Append(new Bold());
                }
                else
                {
                    var newRun = new Run(new Text(valueWithSpace));
                    var newRunProps = new RunProperties();
                    newRunProps.Append(new Bold());
                    newRun.InsertAt(newRunProps, 0);
                    para.Append(newRun);
                }

                return;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error replacing field value: {ex.Message}");
        }
    }

    void ReplaceSignatory(WordprocessingDocument doc, string placeholder, string value, bool isBold)
    {
        var body = doc.MainDocumentPart?.Document?.Body;
        if (body == null) return;

        var allTextElements = body.Descendants<Text>().ToList();
        bool replaced = false;

        for (int i = 0; i < allTextElements.Count; i++)
        {
            if (!replaced && allTextElements[i].Text == placeholder)
            {
                allTextElements[i].Text = value;
                ApplyFormatting(allTextElements[i], isBold);
                
                if (i + 1 < allTextElements.Count && allTextElements[i + 1].Text == ".")
                {
                    allTextElements[i + 1].Text = "";
                }
                
                replaced = true;
                break;
            }
        }

        if (!replaced)
        {
            for (int i = 0; i < allTextElements.Count; i++)
            {
                if (!replaced && allTextElements[i].Text.Contains(placeholder))
                {
                    allTextElements[i].Text = allTextElements[i].Text.Replace(placeholder, value);
                    ApplyFormatting(allTextElements[i], isBold);
                    
                    if (i + 1 < allTextElements.Count && allTextElements[i + 1].Text == ".")
                    {
                        allTextElements[i + 1].Text = "";
                    }
                    
                    replaced = true;
                    break;
                }
            }
        }
    }

    void ApplyFormatting(Text textElem, bool isBold)
    {
        var run = textElem.Ancestors<Run>().FirstOrDefault();
        if (run != null)
        {
            var runProps = run.Elements<RunProperties>().FirstOrDefault();
            if (runProps == null)
            {
                runProps = new RunProperties();
                run.InsertAt(runProps, 0);
            }
            runProps.RemoveAllChildren<Bold>();
            
            if (isBold)
            {
                runProps.Append(new Bold());
            }
        }
    }
    
    private void InitializeComponent()
    {
        this.SuspendLayout();
        this.ResumeLayout(false);
    }
}

// ====================================================================
// LOAD BATCH VIEW
// ====================================================================
public partial class LoadBatchView : UserControl
{
    private Panel? contentPanel;
    private int currentPage = 0;
    private int itemsPerPage = 5;
    private List<string> allBatchFolders = new();
    private string searchQuery = "";
    private Panel? batchesPanel;
    private TextBox? searchBox;
    private Label? pageLabel;
    
    // UI elements to hide when editing
    private Button? btnBrowse;
    private Panel? infoCard;
    private Label? recentLabel;
    private Panel? borderTop;
    private Panel? borderBottom;
    private Panel? borderLeft;
    private Panel? borderRight;
    private Panel? paginationPanel;
    
    // Batch edit form fields
    private Panel? batchEditPanel;
    private NumericUpDown? editNumCandidates;
    private ComboBox? editCboQualification;
    private TextBox? editTxtAssessor;
    private TextBox? editTxtRQMCode;
    private TextBox? editTxtScholarship;
    private TextBox? editTxtTrainingDuration;
    private DateTimePicker? editDtAssessment;
    private string? currentEditingBatchPath;
    
    public LoadBatchView()
    {
        InitializeComponent();
        SetupView();
    }
    
    private void SetupView()
    {
        this.Dock = DockStyle.Fill;
        this.BackColor = Color.FromArgb(241, 245, 249);
        
        Panel topBar = new Panel
        {
            Height = 80,
            Dock = DockStyle.Top,
            BackColor = Color.White,
            Padding = new Padding(30, 15, 30, 15)
        };
        
        Label title = new Label
        {
            Text = "Load Existing Batch",
            Font = new Font("Segoe UI Semibold", 16F, FontStyle.Bold),
            ForeColor = Color.FromArgb(30, 41, 59),
            AutoSize = true,
            Location = new Point(300, 15)
        };
        
        Label subtitle = new Label
        {
            Text = "Select a previously saved batch to continue working",
            Font = new Font("Segoe UI", 9F),
            ForeColor = Color.FromArgb(100, 116, 139),
            AutoSize = true,
            Location = new Point(300, 55)
        };
        
        topBar.Controls.Add(title);
        topBar.Controls.Add(subtitle);
        this.Controls.Add(topBar);
        
        contentPanel = new Panel
        {
            Location = new Point(30, 110),
            Width = 900,
            Height = 600,
            BackColor = Color.FromArgb(241, 245, 249),
            AutoScroll = true,
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
        };
        
        btnBrowse = new Button
        {
            Text = "📂 Browse for Batch Folder",
            Location = new Point(300, 0),
            Width = 900,
            Height = 50,
            BackColor = Color.FromArgb(37, 99, 235),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI Semibold", 11F, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnBrowse.FlatAppearance.BorderSize = 0;
        btnBrowse.Click += BrowseForBatch;
        contentPanel.Controls.Add(btnBrowse);
        
        infoCard = new Panel
        {
            Location = new Point(300, 70),
            Width = 900,
            Height = 150,
            BackColor = Color.FromArgb(239, 246, 255),
            BorderStyle = BorderStyle.FixedSingle
        };
        
        Label infoIcon = new Label
        {
            Text = "ℹ",
            Font = new Font("Segoe UI", 24F),
            ForeColor = Color.FromArgb(37, 99, 235),
            Location = new Point(15, 50),
            AutoSize = true
        };
        
        Label infoTitle = new Label
        {
            Text = "How to Load a Batch",
            Font = new Font("Segoe UI Semibold", 11F, FontStyle.Bold),
            ForeColor = Color.FromArgb(30, 41, 59),
            Location = new Point(110, 20),
            AutoSize = true
        };
        
        Label infoText = new Label
        {
            Text = "Click the 'Browse for Batch Folder' button above and select a folder containing:\n" +
                   "• Billing_AssessorFee.docx\n" +
                   "• Billing_AssessmentFee.docx\n" +
                   "• Attendance_Sheet.docx",
            Font = new Font("Segoe UI", 9F),
            ForeColor = Color.FromArgb(71, 85, 105),
            Location = new Point(110, 50),
            Width = 800,
            Height = 80,
            AutoSize = false
        };
        
        infoCard.Controls.Add(infoIcon);
        infoCard.Controls.Add(infoTitle);
        infoCard.Controls.Add(infoText);
        contentPanel.Controls.Add(infoCard);
        
        // Search Bar
        searchBox = new TextBox
        {
            Location = new Point(300, 230),
            Width = 300,
            Height = 35,
            Font = new Font("Segoe UI", 10F),
            PlaceholderText = "Search batches..."
        };
        searchBox.TextChanged += (s, e) => 
        {
            searchQuery = searchBox.Text;
            currentPage = 0;
            RefreshBatchList();
        };
        contentPanel.Controls.Add(searchBox);
        
        recentLabel = new Label
        {
            Text = "All Batches",
            Font = new Font("Segoe UI Semibold", 12F, FontStyle.Bold),
            ForeColor = Color.FromArgb(30, 41, 59),
            Location = new Point(0, 140),
            AutoSize = true
        };
        contentPanel.Controls.Add(recentLabel);
        
        // Border panels - adjust Width/Height values to change border thickness
        borderTop = new Panel { Location = new Point(270, 290), Width = 1030, Height = 2, BackColor = Color.FromArgb(100, 116, 139) };
        borderBottom = new Panel { Location = new Point(270, 660), Width = 1030, Height = 2, BackColor = Color.FromArgb(100, 116, 139) };
        borderLeft = new Panel { Location = new Point(270, 290), Width = 2, Height = 370, BackColor = Color.FromArgb(100, 116, 139) };
        borderRight = new Panel { Location = new Point(1298, 290), Width = 2, Height = 370, BackColor = Color.FromArgb(100, 116, 139) };
        
        contentPanel.Controls.Add(borderTop);
        contentPanel.Controls.Add(borderBottom);
        contentPanel.Controls.Add(borderLeft);
        contentPanel.Controls.Add(borderRight);
        
        batchesPanel = new Panel
        {
            Location = new Point(0, 300),
            Width = 1300,
            Height = 350,
            AutoScroll = true,
            BackColor = Color.FromArgb(241, 245, 249)
        };
        contentPanel.Controls.Add(batchesPanel);
        
        // Pagination controls
        paginationPanel = new Panel
        {
            Location = new Point(300, 680),
            Width = 880,
            Height = 40,
            BackColor = Color.FromArgb(241, 245, 249)
        };
        
        Button btnPrevious = new RoundedButton
        {
            Text = "← Previous",
            Location = new Point(300, 0),
            Width = 100,
            Height = 35,
            BackColor = Color.FromArgb(100, 116, 139),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9F),
            Cursor = Cursors.Hand,
            BorderRadius = 10
        };
        btnPrevious.FlatAppearance.BorderSize = 0;
        btnPrevious.Click += (s, e) =>
        {
            if (currentPage > 0)
            {
                currentPage--;
                RefreshBatchList();
                if (pageLabel != null)
                    pageLabel.Text = $"Page {currentPage + 1}";
            }
        };
        paginationPanel.Controls.Add(btnPrevious);
        
        pageLabel = new Label
        {
            Text = "Page 1",
            Location = new Point(420, 8),
            Width = 60,
            Height = 20,
            AutoSize = false,
            TextAlign = ContentAlignment.MiddleCenter,
            Font = new Font("Segoe UI", 9F)
        };
        paginationPanel.Controls.Add(pageLabel);
        
        Button btnNext = new RoundedButton
        {
            Text = "Next →",
            Location = new Point(500, 0),
            Width = 100,
            Height = 35,
            BackColor = Color.FromArgb(37, 99, 235),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9F),
            Cursor = Cursors.Hand,
            BorderRadius = 10
        };
        btnNext.FlatAppearance.BorderSize = 0;
        btnNext.Click += (s, e) =>
        {
            var filtered = GetFilteredBatches();
            if ((currentPage + 1) * itemsPerPage < filtered.Count)
            {
                currentPage++;
                RefreshBatchList();
                if (pageLabel != null)
                    pageLabel.Text = $"Page {currentPage + 1}";
            }
        };
        paginationPanel.Controls.Add(btnNext);
        
        contentPanel.Controls.Add(paginationPanel);
        
        this.Controls.Add(contentPanel);
        
        RefreshBatchList();
    }
    
    private void RefreshBatchList()
    {
        if (batchesPanel == null)
            return;
        batchesPanel.Controls.Clear();
        
        string outputDir = "Output";
        if (!Directory.Exists(outputDir))
        {
            Directory.CreateDirectory(outputDir);
            Label noBatches = new Label
            {
                Text = "No batches found. Create your first batch in the 'New Batch' tab!",
                Font = new Font("Segoe UI", 10F, FontStyle.Italic),
                ForeColor = Color.FromArgb(100, 116, 139),
                Location = new Point(0, 150),
                Width = 1300,
                Height = 50,
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleCenter
            };
            batchesPanel.Controls.Add(noBatches);
            return;
        }
        
        allBatchFolders = Directory.GetDirectories(outputDir)
            .Where(f => File.Exists(Path.Combine(f, "batch_info.json")))
            .OrderByDescending(f => Directory.GetLastWriteTime(f))
            .ToList();
        
        var filtered = GetFilteredBatches();
        
        if (filtered.Count == 0)
        {
            Label noBatches = new Label
            {
                Text = "No batches found. Create your first batch in the 'New Batch' tab!",
                Font = new Font("Segoe UI", 10F, FontStyle.Italic),
                ForeColor = Color.FromArgb(100, 116, 139),
                Location = new Point(0, 20),
                Width = 1300,
                Height = 50,
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleCenter
            };
            batchesPanel.Controls.Add(noBatches);
            return;
        }
        
        int startIdx = currentPage * itemsPerPage;
        int endIdx = Math.Min(startIdx + itemsPerPage, filtered.Count);
        
        int yPos =  0;
        for (int i = startIdx; i < endIdx; i++)
        {
            try
            {
                var folder = filtered[i];
                var batchInfo = LoadBatchInfoJson(folder);
                if (batchInfo != null)
                {
                    AddBatchItem(folder, batchInfo, yPos);
                    yPos += 130;
                }
            }
            catch { }
        }
    }
    
    private List<string> GetFilteredBatches()
    {
        if (string.IsNullOrWhiteSpace(searchQuery))
            return allBatchFolders;
        
        return allBatchFolders
            .Where(f => Path.GetFileName(f).Contains(searchQuery, StringComparison.OrdinalIgnoreCase))
            .ToList();
    }
    
    private void LoadRecentBatches(int startY)
    {
        RefreshBatchList();
    }
    
    private BatchInfo? LoadBatchInfoJson(string folderPath)
    {
        try
        {
            string jsonPath = Path.Combine(folderPath, "batch_info.json");
            if (File.Exists(jsonPath))
            {
                string json = File.ReadAllText(jsonPath);
                var result = JsonSerializer.Deserialize<BatchInfo>(json);
                return result;
            }
        }
        catch { }
        return null;
    }
    
    private void AddBatchItem(string folderPath, BatchInfo batchInfo, int yPos)
    {
        Panel item = new Panel
        {
            Location = new Point(300, yPos),
            Width = 880,
            Height = 120,
            BackColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle,
            Cursor = Cursors.Hand
        };
        
        string batchName = Path.GetFileName(folderPath);
        Label lblBatchName = new Label
        {
            Text = batchName,
            Font = new Font("Segoe UI Semibold", 11F, FontStyle.Bold),
            ForeColor = Color.FromArgb(30, 41, 59),
            Location = new Point(15, 10),
            Width = 700,
            Height = 25,
            AutoSize = false
        };
        
        Label lblQualification = new Label
        {
            Text = $"📄 {batchInfo.Qualification}",
            Location = new Point(15, 40),
            Width = 500,
            AutoSize = false,
            Font = new Font("Segoe UI", 9F),
            ForeColor = Color.FromArgb(71, 85, 105)
        };
        
        Label lblCandidates = new Label
        {
            Text = $"👥 {batchInfo.CandidateCount} Candidates",
            Location = new Point(430, 40),
            Width = 200,
            AutoSize = false,
            Font = new Font("Segoe UI", 9F),
            ForeColor = Color.FromArgb(71, 85, 105)
        };
        
        Label lblAssessor = new Label
        {
            Text = $"Assessor: {batchInfo.Assessor}",
            Location = new Point(15, 65),
            Width = 500,
            AutoSize = false,
            Font = new Font("Segoe UI", 9F),
            ForeColor = Color.FromArgb(71, 85, 105)
        };
        
        Label lblDate = new Label
        {
            Text = $"📅 {batchInfo.AssessmentDate}",
            Location = new Point(530, 65),
            Width = 200,
            AutoSize = false,
            Font = new Font("Segoe UI", 9F),
            ForeColor = Color.FromArgb(71, 85, 105)
        };
        
        Button btnLoad = new Button
        {
            Text = "Open Folder",
            Location = new Point(730, 75),
            Width = 130,
            Height = 35,
            BackColor = Color.FromArgb(100, 116, 139),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold),
            Cursor = Cursors.Hand,
            Tag = folderPath
        };
        btnLoad.FlatAppearance.BorderSize = 0;
        btnLoad.Click += (s, e) => 
        {
            try
            {
                System.Diagnostics.Process.Start("explorer.exe", folderPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error opening folder: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        };
        
        Button btnLoadData = new Button
        {
            Text = "Load Batch",
            Location = new Point(630, 75),
            Width = 90,
            Height = 35,
            BackColor = Color.FromArgb(22, 163, 74),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold),
            Cursor = Cursors.Hand,
            Tag = folderPath
        };
        btnLoadData.FlatAppearance.BorderSize = 0;
        btnLoadData.Click += (s, e) => LoadBatchData(folderPath, batchInfo);
        
        item.Controls.Add(lblBatchName);
        item.Controls.Add(lblQualification);
        item.Controls.Add(lblCandidates);
        item.Controls.Add(lblAssessor);
        item.Controls.Add(lblDate);
        item.Controls.Add(btnLoadData);
        item.Controls.Add(btnLoad);
        
        if (batchesPanel != null)
            batchesPanel.Controls.Add(item);
    }
    
    private void LoadBatchData(string folderPath, BatchInfo batchInfo)
    {
        try
        {
            currentEditingBatchPath = folderPath;
            ShowBatchEditForm(batchInfo);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error loading batch: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
    
    private void ShowBatchEditForm(BatchInfo batchInfo)
    {
        // Hide all previous UI elements
        if (btnBrowse != null) btnBrowse.Visible = false;
        if (infoCard != null) infoCard.Visible = false;
        if (searchBox != null) searchBox.Visible = false;
        if (recentLabel != null) recentLabel.Visible = false;
        if (borderTop != null) borderTop.Visible = false;
        if (borderBottom != null) borderBottom.Visible = false;
        if (borderLeft != null) borderLeft.Visible = false;
        if (borderRight != null) borderRight.Visible = false;
        if (batchesPanel != null) batchesPanel.Visible = false;
        if (paginationPanel != null) paginationPanel.Visible = false;
        
        if (batchEditPanel != null)
        {
            batchEditPanel.Visible = true;
        }
        else
        {
            // Create the edit form panel - centered
            batchEditPanel = new Panel
            {
                Location = new Point(50, 150),
                Width = 1000,
                Height = 600,
                BackColor = Color.FromArgb(241, 245, 249),
                AutoScroll = true
            };
            
            // Title
            Label editTitle = new Label
            {
                Text = "Edit Batch",
                Font = new Font("Segoe UI Semibold", 14F, FontStyle.Bold),
                ForeColor = Color.FromArgb(30, 41, 59),
                Location = new Point(50, 10),
                AutoSize = true
            };
            batchEditPanel.Controls.Add(editTitle);
            
            // Candidates card
            Panel candCard = CreateEditCard("Candidates Information", 50, 50);
            Label lblCandidates = new Label { Text = "Number of Candidates (Max 25)", Location = new Point(20, 50), Width = 300, Height = 25, AutoSize = false, Font = new Font("Segoe UI", 9F) };
            editNumCandidates = new NumericUpDown { Location = new Point(20, 70), Width = 150, Maximum = 25, Minimum = 0, Font = new Font("Segoe UI", 10F) };
            candCard.Controls.Add(lblCandidates);
            candCard.Controls.Add(editNumCandidates);
            batchEditPanel.Controls.Add(candCard);
            
            // Qualification card
            Panel qualCard = CreateEditCard("Qualification & Assessor", 50, 180);
            Label lblQual = new Label { Text = "Qualification *", Location = new Point(20, 50), AutoSize = true, Font = new Font("Segoe UI", 9F) };
            editCboQualification = new ComboBox { Location = new Point(20, 70), Width = 710, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Segoe UI", 10F) };
            editCboQualification.Items.AddRange(new object[]
            {
                "Barangay Health Services NC II", "Bartending NC II", "Bookkeeping NC II",
                "Bread and Pastry Production NC II", "Computer System Servicing NC II", "Cookery NC II",
                "Driving NC II", "Electrical Installation and Maintenance NC II", "Event Management Services NC III",
                "Food and Beverage Services NC II", "Housekeeping NC II", "Motorcycle/Small Engine Servicing NC II"
            });
            Label lblAssessor = new Label { Text = "Assessor Name *", Location = new Point(20, 110), AutoSize = true, Font = new Font("Segoe UI", 9F) };
            editTxtAssessor = new TextBox { Location = new Point(20, 130), Width = 710, Font = new Font("Segoe UI", 10F) };
            qualCard.Controls.Add(lblQual);
            qualCard.Controls.Add(editCboQualification);
            qualCard.Controls.Add(lblAssessor);
            qualCard.Controls.Add(editTxtAssessor);
            batchEditPanel.Controls.Add(qualCard);
            
            // Admin card
            Panel adminCard = CreateEditCard("Administrative Details", 50, 350);
            Label lblRQM = new Label { Text = "RQM Code *", Location = new Point(20, 50), AutoSize = true, Font = new Font("Segoe UI", 9F) };
            editTxtRQMCode = new TextBox { Location = new Point(20, 70), Width = 340, Font = new Font("Segoe UI", 10F) };
            Label lblSchol = new Label { Text = "Type of Scholarship *", Location = new Point(390, 50), AutoSize = true, Font = new Font("Segoe UI", 9F) };
            editTxtScholarship = new TextBox { Location = new Point(390, 70), Width = 340, Font = new Font("Segoe UI", 10F) };
            adminCard.Controls.Add(lblRQM);
            adminCard.Controls.Add(editTxtRQMCode);
            adminCard.Controls.Add(lblSchol);
            adminCard.Controls.Add(editTxtScholarship);
            batchEditPanel.Controls.Add(adminCard);
            
            // Training card
            Panel trainingCard = CreateEditCard("Training & Assessment", 50, 480);
            Label lblDuration = new Label { Text = "Training Duration *", Location = new Point(20, 50), AutoSize = true, Font = new Font("Segoe UI", 9F) };
            editTxtTrainingDuration = new TextBox { Location = new Point(20, 70), Width = 340, Font = new Font("Segoe UI", 10F) };
            Label lblDate = new Label { Text = "Assessment Date *", Location = new Point(390, 50), AutoSize = true, Font = new Font("Segoe UI", 9F) };
            editDtAssessment = new DateTimePicker { Location = new Point(390, 70), Width = 340, Font = new Font("Segoe UI", 10F) };
            trainingCard.Controls.Add(lblDuration);
            trainingCard.Controls.Add(editTxtTrainingDuration);
            trainingCard.Controls.Add(lblDate);
            trainingCard.Controls.Add(editDtAssessment);
            batchEditPanel.Controls.Add(trainingCard);
            
            // Buttons
            Button btnSaveEdit = new Button
            {
                Text = "Save Edit",
                Location = new Point(70, 520),
                Width = 150,
                Height = 40,
                BackColor = Color.FromArgb(22, 163, 74),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI Semibold", 10F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnSaveEdit.FlatAppearance.BorderSize = 0;
            btnSaveEdit.Click += (s, e) => SaveBatchEdit();
            batchEditPanel.Controls.Add(btnSaveEdit);
            
            Button btnCancel = new Button
            {
                Text = "Cancel",
                Location = new Point(230, 520),
                Width = 100,
                Height = 40,
                BackColor = Color.FromArgb(100, 116, 139),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnCancel.FlatAppearance.BorderSize = 0;
            btnCancel.Click += (s, e) => CancelBatchEdit();
            batchEditPanel.Controls.Add(btnCancel);
            
            if (contentPanel != null)
                contentPanel.Controls.Add(batchEditPanel);
        }
        
        // Populate fields
        if (editNumCandidates != null) editNumCandidates.Value = batchInfo.CandidateCount;
        if (editCboQualification != null) editCboQualification.SelectedItem = batchInfo.Qualification;
        if (editTxtAssessor != null) editTxtAssessor.Text = batchInfo.Assessor;
        if (editTxtRQMCode != null) editTxtRQMCode.Text = batchInfo.RQMCode;
        if (editTxtScholarship != null) editTxtScholarship.Text = batchInfo.Scholarship;
        if (editTxtTrainingDuration != null) editTxtTrainingDuration.Text = batchInfo.TrainingDuration;
        if (editDtAssessment != null && DateTime.TryParse(batchInfo.AssessmentDate, out DateTime date))
            editDtAssessment.Value = date;
    }
    
    private Panel CreateEditCard(string title, int x, int y)
    {
        Panel card = new Panel
        {
            Location = new Point(x, y),
            Width = 750,
            Height = 150,
            BackColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle
        };
        
        Panel header = new Panel { Height = 40, Dock = DockStyle.Top, BackColor = Color.FromArgb(248, 250, 252) };
        Label headerLabel = new Label { Text = title, Font = new Font("Segoe UI Semibold", 10F, FontStyle.Bold), ForeColor = Color.FromArgb(30, 41, 59), Location = new Point(15, 10), AutoSize = true };
        header.Controls.Add(headerLabel);
        card.Controls.Add(header);
        
        return card;
    }
    
    private void SaveBatchEdit()
    {
        MessageBox.Show("Batch saved successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        CancelBatchEdit();
    }
    
    private void CancelBatchEdit()
    {
        // Show all UI elements again
        if (btnBrowse != null) btnBrowse.Visible = true;
        if (infoCard != null) infoCard.Visible = true;
        if (searchBox != null) searchBox.Visible = true;
        if (recentLabel != null) recentLabel.Visible = true;
        if (borderTop != null) borderTop.Visible = true;
        if (borderBottom != null) borderBottom.Visible = true;
        if (borderLeft != null) borderLeft.Visible = true;
        if (borderRight != null) borderRight.Visible = true;
        if (batchesPanel != null) batchesPanel.Visible = true;
        if (paginationPanel != null) paginationPanel.Visible = true;
        
        if (batchEditPanel != null)
            batchEditPanel.Visible = false;
        
        RefreshBatchList();
    }
    
    private void BrowseForBatch(object? sender, EventArgs e)
    {
        using (var folderDialog = new FolderBrowserDialog())
        {
            folderDialog.Description = "Select the folder containing your generated forms (with 3 .docx files)";
            
            if (folderDialog.ShowDialog() == DialogResult.OK)
            {
                string selectedFolder = folderDialog.SelectedPath;
                
                try
                {
                    System.Diagnostics.Process.Start("explorer.exe", selectedFolder);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error opening folder: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }
    
    private void InitializeComponent()
    {
        this.SuspendLayout();
        this.ResumeLayout(false);
    }
}

// ====================================================================
// REPORTS VIEW
// ====================================================================
public partial class ReportsView : UserControl
{
    public ReportsView()
    {
        InitializeComponent();
        SetupView();
    }
    
    private void SetupView()
    {
        this.Dock = DockStyle.Fill;
        this.BackColor = Color.FromArgb(241, 245, 249);
        this.AutoScroll = true;
        
        Panel topBar = new Panel
        {
            Height = 80,
            Dock = DockStyle.Top,
            BackColor = Color.White,
            Padding = new Padding(30, 15, 30, 15)
        };
        
        Label title = new Label
        {
            Text = "Reports & Analytics",
            Font = new Font("Segoe UI Semibold", 16F, FontStyle.Bold),
            ForeColor = Color.FromArgb(30, 41, 59),
            AutoSize = true,
            Location = new Point(30, 15)
        };
        
        Label subtitle = new Label
        {
            Text = "View batch statistics and export data",
            Font = new Font("Segoe UI", 9F),
            ForeColor = Color.FromArgb(100, 116, 139),
            AutoSize = true,
            Location = new Point(30, 45)
        };
        
        topBar.Controls.Add(title);
        topBar.Controls.Add(subtitle);
        this.Controls.Add(topBar);
        
        Panel content = new Panel
        {
            Location = new Point(0, 80),
            Width = this.Width,
            Height = this.Height - 80,
            BackColor = Color.FromArgb(241, 245, 249),
            AutoScroll = true,
            Padding = new Padding(30),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
        };
        
        CreateStatisticsCards(content, 30, 30);
        
        Panel actionsCard = CreateQuickActionsCard();
        actionsCard.Location = new Point(30, 160);
        content.Controls.Add(actionsCard);
        
        this.Controls.Add(content);
    }
    
    private void CreateStatisticsCards(Panel parent, int x, int y)
    {
        int totalBatches = CountBatches();
        int totalCandidates = CountTotalCandidates();
        
        CreateStatCard(parent, "Total Batches", totalBatches.ToString(), x, y, Color.FromArgb(37, 99, 235));
        CreateStatCard(parent, "Total Candidates", totalCandidates.ToString(), x + 230, y, Color.FromArgb(22, 163, 74));
        CreateStatCard(parent, "This Month", "0", x + 460, y, Color.FromArgb(168, 85, 247));
        CreateStatCard(parent, "Documents", (totalBatches * 3).ToString(), x + 690, y, Color.FromArgb(249, 115, 22));
    }
    
    private void CreateStatCard(Panel parent, string label, string value, int x, int y, Color accentColor)
    {
        Panel card = new Panel
        {
            Location = new Point(x, y),
            Width = 210,
            Height = 110,
            BackColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle
        };
        
        Label lblLabel = new Label
        {
            Text = label,
            Location = new Point(15, 15),
            AutoSize = true,
            Font = new Font("Segoe UI", 8.5F),
            ForeColor = Color.FromArgb(100, 116, 139)
        };
        
        Label lblValue = new Label
        {
            Text = value,
            Location = new Point(15, 45),
            AutoSize = true,
            Font = new Font("Segoe UI", 22F, FontStyle.Bold),
            ForeColor = Color.FromArgb(30, 41, 59)
        };
        
        Panel colorBar = new Panel
        {
            Location = new Point(0, 0),
            Width = 4,
            Height = 110,
            BackColor = accentColor
        };
        
        card.Controls.Add(colorBar);
        card.Controls.Add(lblLabel);
        card.Controls.Add(lblValue);
        parent.Controls.Add(card);
    }
    
    private Panel CreateQuickActionsCard()
    {
        Panel card = new Panel
        {
            Width = 900,
            Height = 250,
            BackColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle,
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        };
        
        Panel header = new Panel
        {
            Height = 50,
            Dock = DockStyle.Top,
            BackColor = Color.FromArgb(248, 250, 252)
        };
        
        Label headerLabel = new Label
        {
            Text = "Quick Actions",
            Font = new Font("Segoe UI Semibold", 10F, FontStyle.Bold),
            ForeColor = Color.FromArgb(30, 41, 59),
            Location = new Point(20, 15),
            AutoSize = true
        };
        header.Controls.Add(headerLabel);
        card.Controls.Add(header);
        
        AddActionItem(card, "📂 Open Output Folder", "View all generated documents", 60, OpenOutputFolder);
        AddActionItem(card, "📊 Refresh Statistics", "Update the statistics above", 130, RefreshStats);
        
        return card;
    }
    
    private void AddActionItem(Panel parent, string title, string description, int yPos, EventHandler clickHandler)
    {
        Panel item = new Panel
        {
            Location = new Point(20, yPos),
            Width = 710,
            Height = 60,
            BackColor = Color.FromArgb(249, 250, 251),
            BorderStyle = BorderStyle.FixedSingle
        };
        
        Label lblTitle = new Label
        {
            Text = title,
            Location = new Point(15, 8),
            Width = 650,
            AutoSize = false,
            Font = new Font("Segoe UI Semibold", 10F, FontStyle.Bold),
            ForeColor = Color.FromArgb(30, 41, 59)
        };
        
        Label lblDesc = new Label
        {
            Text = description,
            Location = new Point(15, 30),
            Width = 650,
            AutoSize = false,
            Font = new Font("Segoe UI", 8.5F),
            ForeColor = Color.FromArgb(100, 116, 139)
        };
        
        Button btnAction = new Button
        {
            Text = "Open",
            Location = new Point(740, 13),
            Width = 100,
            Height = 32,
            BackColor = Color.FromArgb(37, 99, 235),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnAction.FlatAppearance.BorderSize = 0;
        btnAction.Click += clickHandler;
        
        item.Controls.Add(lblTitle);
        item.Controls.Add(lblDesc);
        item.Controls.Add(btnAction);
        parent.Controls.Add(item);
    }
    
    private int CountBatches()
    {
        try
        {
            string outputDir = "Output";
            if (!Directory.Exists(outputDir))
                return 0;
            
            return Directory.GetDirectories(outputDir)
                .Count(f => File.Exists(Path.Combine(f, "batch_info.json")));
        }
        catch
        {
            return 0;
        }
    }
    
    private int CountTotalCandidates()
    {
        try
        {
            string outputDir = "Output";
            if (!Directory.Exists(outputDir))
                return 0;
            
            int total = 0;
            var folders = Directory.GetDirectories(outputDir)
                .Where(f => File.Exists(Path.Combine(f, "batch_info.json")));
            
            foreach (var folder in folders)
            {
                try
                {
                    string jsonPath = Path.Combine(folder, "batch_info.json");
                    string json = File.ReadAllText(jsonPath);
                    var batchInfo = JsonSerializer.Deserialize<BatchInfo>(json);
                    if (batchInfo != null)
                        total += batchInfo.CandidateCount;
                }
                catch { }
            }
            
            return total;
        }
        catch
        {
            return 0;
        }
    }
    
    private void OpenOutputFolder(object? sender, EventArgs e)
    {
        try
        {
            string outputDir = "Output";
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);
            
            System.Diagnostics.Process.Start("explorer.exe", outputDir);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error opening folder: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
    
    private void RefreshStats(object? sender, EventArgs e)
    {
        this.Controls.Clear();
        SetupView();
        MessageBox.Show("Statistics refreshed!", "Refresh", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }
    
    private void InitializeComponent()
    {
        this.SuspendLayout();
        this.ResumeLayout(false);
    }
}

// ====================================================================
// CANDIDATE FORM (DIALOG)
// ====================================================================
public class CandidateForm : Form
{
    List<TextBox> txtNames = new List<TextBox>();
    List<TextBox> txtReferences = new List<TextBox>();
    List<ComboBox> cboFees = new List<ComboBox>();
    List<TextBox> txtAssessorFees = new List<TextBox>();
    Button btnSave = new Button();
    Button btnCancel = new Button();
    int candidateCount = 0;
    string qualification = "";

    public CandidateForm(int count, List<Candidate> existingCandidates, string qual)
    {
        candidateCount = count;
        qualification = qual;
        Text = $"Enter Candidate Details ({count} candidates)";
        Width = 900;
        Height = 15 + 35 + 25 + (count * 30) + 15 + 35 + 50 + 35 + 20;
        StartPosition = FormStartPosition.CenterParent;
        AutoScroll = false;
        BackColor = Color.FromArgb(245, 245, 250);
        Font = new Font("Segoe UI", 10);

        int y = 15;

        var lblHeader = new Label()
        {
            Text = "📝 Enter Candidate Details",
            Top = y,
            Left = 15,
            Width = 850,
            Height = 25,
            Font = new Font("Segoe UI", 11, FontStyle.Bold),
            ForeColor = Color.FromArgb(0, 102, 204)
        };
        Controls.Add(lblHeader);
        y += 35;

        var lblNoHeader = new Label() { Text = "No.", Top = y, Left = 15, Width = 40, Font = new Font("Segoe UI", 9, FontStyle.Bold) };
        var lblNameHeader = new Label() { Text = "Candidate Name", Top = y, Left = 60, Width = 220, Font = new Font("Segoe UI", 9, FontStyle.Bold) };
        var lblRefHeader = new Label() { Text = "Reference Number", Top = y, Left = 290, Width = 180, Font = new Font("Segoe UI", 9, FontStyle.Bold) };
        var lblAssessmentFeeHeader = new Label() { Text = "Assessment Fee", Top = y, Left = 480, Width = 140, Font = new Font("Segoe UI", 9, FontStyle.Bold) };
        var lblAssessorFeeHeader = new Label() { Text = "Assessor's Fee", Top = y, Left = 630, Width = 140, Font = new Font("Segoe UI", 9, FontStyle.Bold) };
        Controls.Add(lblNoHeader);
        Controls.Add(lblNameHeader);
        Controls.Add(lblRefHeader);
        Controls.Add(lblAssessmentFeeHeader);
        Controls.Add(lblAssessorFeeHeader);
        y += 25;

        var feeOptions = AssessmentFees.GetFeeOptions(qualification);

        for (int i = 0; i < count; i++)
        {
            var lblNo = new Label() { Text = (i + 1).ToString(), Top = y, Left = 15, Width = 40, Font = new Font("Segoe UI", 10) };
            Controls.Add(lblNo);

            var txtName = new TextBox() { Top = y, Left = 60, Width = 220, Height = 24, Font = new Font("Segoe UI", 10) };
            var txtRef = new TextBox() { Top = y, Left = 290, Width = 180, Height = 24, Font = new Font("Segoe UI", 10) };
            var cboFee = new ComboBox() { Top = y, Left = 480, Width = 140, Height = 24, Font = new Font("Segoe UI", 10), DropDownStyle = ComboBoxStyle.DropDownList };
            var txtAssessorFee = new TextBox() { Top = y, Left = 630, Width = 140, Height = 24, Font = new Font("Segoe UI", 10) };

            foreach (var fee in feeOptions)
                cboFee.Items.Add(fee);

            if (i < existingCandidates.Count)
            {
                txtName.Text = existingCandidates[i].Name;
                txtRef.Text = existingCandidates[i].Reference;
                
                if (cboFee.Items.Contains(existingCandidates[i].AssessmentFee))
                    cboFee.SelectedItem = existingCandidates[i].AssessmentFee;
                
                txtAssessorFee.Text = existingCandidates[i].AssessorFee;
            }
            else
            {
                txtName.PlaceholderText = $"Candidate {i + 1}";
                txtRef.PlaceholderText = $"AC-{(i + 1):000}";
                txtAssessorFee.PlaceholderText = "0.00";
            }

            txtName.ForeColor = Color.FromArgb(64, 64, 64);
            txtRef.ForeColor = Color.FromArgb(64, 64, 64);
            txtAssessorFee.ForeColor = Color.FromArgb(64, 64, 64);

            txtNames.Add(txtName);
            txtReferences.Add(txtRef);
            cboFees.Add(cboFee);
            txtAssessorFees.Add(txtAssessorFee);
            Controls.Add(txtName);
            Controls.Add(txtRef);
            Controls.Add(cboFee);
            Controls.Add(txtAssessorFee);

            y += 30;
        }

        y += 15;

        var btnClearAllFees = new Button()
        {
            Text = "🗑 Clear All Fees",
            Top = y,
            Left = 120,
            Width = 160,
            Height = 35,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            BackColor = Color.FromArgb(220, 53, 69),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        btnClearAllFees.FlatAppearance.BorderSize = 0;
        btnClearAllFees.Click += (s, e) =>
        {
            foreach (var cbo in cboFees)
            {
                cbo.SelectedIndex = -1;
                cbo.Text = "";
            }
        };
        Controls.Add(btnClearAllFees);
        y += 50;

        btnSave.Text = "✓ Save";
        btnSave.Top = y;
        btnSave.Left = 300;
        btnSave.Width = 150;
        btnSave.Height = 35;
        btnSave.Font = new Font("Segoe UI", 10, FontStyle.Bold);
        btnSave.BackColor = Color.FromArgb(40, 167, 69);
        btnSave.ForeColor = Color.White;
        btnSave.FlatStyle = FlatStyle.Flat;
        btnSave.FlatAppearance.BorderSize = 0;
        btnSave.Click += (s, e) => { DialogResult = DialogResult.OK; Close(); };
        Controls.Add(btnSave);

        btnCancel.Text = "✕ Cancel";
        btnCancel.Top = y;
        btnCancel.Left = 460;
        btnCancel.Width = 150;
        btnCancel.Height = 35;
        btnCancel.Font = new Font("Segoe UI", 10, FontStyle.Bold);
        btnCancel.BackColor = Color.FromArgb(220, 53, 69);
        btnCancel.ForeColor = Color.White;
        btnCancel.FlatStyle = FlatStyle.Flat;
        btnCancel.FlatAppearance.BorderSize = 0;
        btnCancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };
        Controls.Add(btnCancel);
    }

    public List<Candidate> GetCandidates()
    {
        var candidates = new List<Candidate>();
        for (int i = 0; i < candidateCount; i++)
        {
            if (i < txtNames.Count && i < txtReferences.Count && i < cboFees.Count && i < txtAssessorFees.Count)
            {
                candidates.Add(new Candidate()
                {
                    Name = txtNames[i].Text.Trim(),
                    Reference = txtReferences[i].Text.Trim(),
                    AssessmentFee = cboFees[i].SelectedItem?.ToString() ?? "",
                    AssessorFee = txtAssessorFees[i].Text.Trim()
                });
            }
        }
        return candidates;
    }
}

// ====================================================================
// CUSTOM DATES FORM (DIALOG)
// ====================================================================
public class CustomDatesForm : Form
{
    List<DateTimePicker> dtPickers = new List<DateTimePicker>();
    Button btnSave = new Button();
    Button btnCancel = new Button();
    int dateCount = 0;

    public CustomDatesForm(int count, DateTime firstDate, List<DateTime> existingDates)
    {
        dateCount = count;
        Text = $"Attendance Sheet - Page Dates ({count} additional date{(count > 1 ? "s" : "")})";
        Width = 500;
        Height = Math.Max(280, 200 + (count * 60));
        StartPosition = FormStartPosition.CenterParent;
        AutoScroll = false;
        BackColor = Color.FromArgb(245, 245, 250);
        Font = new Font("Segoe UI", 10);

        int y = 15;

        var lblHeader = new Label()
        {
            Text = "📅 Set Assessment Dates for Additional Pages",
            Top = y,
            Left = 15,
            Width = 450,
            Height = 25,
            Font = new Font("Segoe UI", 11, FontStyle.Bold),
            ForeColor = Color.FromArgb(0, 102, 204)
        };
        Controls.Add(lblHeader);
        y += 35;

        var lblInfo = new Label()
        {
            Text = $"First date (Page 1): {firstDate:MM/dd/yyyy} [This is your main assessment date]",
            Top = y,
            Left = 15,
            Width = 450,
            Height = 20,
            Font = new Font("Segoe UI", 9, FontStyle.Italic),
            ForeColor = Color.FromArgb(100, 100, 100)
        };
        Controls.Add(lblInfo);
        y += 30;

        for (int i = 0; i < count; i++)
        {
            var lblDate = new Label()
            {
                Text = $"Page {i + 2} Date:",
                Top = y,
                Left = 15,
                Width = 100,
                Font = new Font("Segoe UI", 10)
            };
            Controls.Add(lblDate);

            var dtPicker = new DateTimePicker()
            {
                Top = y,
                Left = 120,
                Width = 200,
                Font = new Font("Segoe UI", 10)
            };

            if (i < existingDates.Count && existingDates[i] != DateTime.MinValue)
            {
                dtPicker.Value = existingDates[i];
            }
            else
            {
                dtPicker.Value = firstDate.AddDays(i + 1);
            }

            dtPickers.Add(dtPicker);
            Controls.Add(dtPicker);

            y += 50;
        }

        y += 15;

        btnSave.Text = "✓ Save Dates";
        btnSave.Top = y;
        btnSave.Left = 120;
        btnSave.Width = 140;
        btnSave.Height = 35;
        btnSave.Font = new Font("Segoe UI", 10, FontStyle.Bold);
        btnSave.BackColor = Color.FromArgb(40, 167, 69);
        btnSave.ForeColor = Color.White;
        btnSave.FlatStyle = FlatStyle.Flat;
        btnSave.FlatAppearance.BorderSize = 0;
        btnSave.Click += (s, e) => { DialogResult = DialogResult.OK; Close(); };
        Controls.Add(btnSave);

        btnCancel.Text = "✕ Cancel";
        btnCancel.Top = y;
        btnCancel.Left = 270;
        btnCancel.Width = 140;
        btnCancel.Height = 35;
        btnCancel.Font = new Font("Segoe UI", 10, FontStyle.Bold);
        btnCancel.BackColor = Color.FromArgb(220, 53, 69);
        btnCancel.ForeColor = Color.White;
        btnCancel.FlatStyle = FlatStyle.Flat;
        btnCancel.FlatAppearance.BorderSize = 0;
        btnCancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };
        Controls.Add(btnCancel);
    }

    public List<DateTime> GetDates()
    {
        var dates = new List<DateTime>();
        foreach (var dt in dtPickers)
        {
            dates.Add(dt.Value);
        }
        return dates;
    }
}

// ====================================================================
// EDIT FORM (DIALOG)
// ====================================================================
public class EditForm : Form
{
    NumericUpDown numCandidates;
    NumericUpDown edtNumCandidates;
    ComboBox cboQualification;
    ComboBox edtCboQualification;
    TextBox txtAssessor;
    TextBox edtTxtAssessor;
    TextBox txtRQMCode;
    TextBox edtTxtRQMCode;
    TextBox txtScholarship;
    TextBox edtTxtScholarship;
    TextBox txtTrainingDuration;
    TextBox edtTxtTrainingDuration;
    DateTimePicker dtAssessment;
    DateTimePicker edtDtAssessment;
    Button btnEditCandidates;
    Button btnEditDates;
    Button btnEditSignatories;
    Button btnDone;
    Button btnCancel;
    List<Candidate> customCandidates;
    List<DateTime> customDates;
    Signatory signatories;

    public EditForm(NumericUpDown numCand, ComboBox cboQual, TextBox txtAss, TextBox txtRQM,
                    TextBox txtSchol, TextBox txtDur, DateTimePicker dtAss,
                    List<Candidate> cands, List<DateTime> dates, Signatory sigs = null!)
    {
        numCandidates = numCand;
        cboQualification = cboQual;
        txtAssessor = txtAss;
        txtRQMCode = txtRQM;
        txtScholarship = txtSchol;
        txtTrainingDuration = txtDur;
        dtAssessment = dtAss;
        customCandidates = new List<Candidate>(cands);
        customDates = new List<DateTime>(dates);
        signatories = sigs ?? new Signatory();

        Text = "Edit Generated Documents";
        Width = 700;
        Height = 700;
        StartPosition = FormStartPosition.CenterParent;
        BackColor = Color.FromArgb(245, 245, 250);
        Font = new Font("Segoe UI", 10);
        AutoScroll = true;

        int y = 20;

        var lblHeader = new Label()
        {
            Text = "📝 Edit Document Information",
            Top = y,
            Left = 20,
            Width = 500,
            Height = 30,
            Font = new Font("Segoe UI", 12, FontStyle.Bold),
            ForeColor = Color.FromArgb(0, 102, 204)
        };
        Controls.Add(lblHeader);
        y += 45;

        var lblCandidates = new Label() { Text = "Number of Candidates", Top = y, Left = 20, Width = 200, Font = new Font("Segoe UI", 10) };
        edtNumCandidates = new NumericUpDown() { Top = y, Left = 230, Width = 80, Maximum = 25, Value = numCandidates.Value };
        Controls.Add(lblCandidates);
        Controls.Add(edtNumCandidates);
        y += 40;

        var lblQual = new Label() { Text = "Qualification", Top = y, Left = 20, Width = 200, Font = new Font("Segoe UI", 10) };
        edtCboQualification = new ComboBox() { Top = y, Left = 230, Width = 430, Height = 28, DropDownStyle = ComboBoxStyle.DropDownList };
        edtCboQualification.Items.AddRange(new object[]
        {
            "Barangay Health Services NC II",
            "Bartending NC II",
            "Bookkeeping NC II",
            "Bread and Pastry Production NC II",
            "Computer System Servicing NC II",
            "Cookery NC II",
            "Driving NC II",
            "Electrical Installation and Maintenance NC II",
            "Event Management Services NC III",
            "Food and Beverage Services NC II",
            "Housekeeping NC II",
            "Motorcycle/Small Engine Servicing NC II"
        });
        edtCboQualification.SelectedIndex = cboQualification.SelectedIndex;
        Controls.Add(lblQual);
        Controls.Add(edtCboQualification);
        y += 40;

        var lblAssessor = new Label() { Text = "Assessor Name", Top = y, Left = 20, Width = 200, Font = new Font("Segoe UI", 10) };
        edtTxtAssessor = new TextBox() { Top = y, Left = 230, Width = 430, Height = 28, Text = txtAssessor.Text };
        Controls.Add(lblAssessor);
        Controls.Add(edtTxtAssessor);
        y += 40;

        var lblRQM = new Label() { Text = "RQM Code", Top = y, Left = 20, Width = 200, Font = new Font("Segoe UI", 10) };
        edtTxtRQMCode = new TextBox() { Top = y, Left = 230, Width = 430, Height = 28, Text = txtRQMCode.Text };
        Controls.Add(lblRQM);
        Controls.Add(edtTxtRQMCode);
        y += 40;

        var lblSchol = new Label() { Text = "Type of Scholarship/Modality", Top = y, Left = 20, Width = 200, Font = new Font("Segoe UI", 10) };
        edtTxtScholarship = new TextBox() { Top = y, Left = 230, Width = 430, Height = 28, Text = txtScholarship.Text };
        Controls.Add(lblSchol);
        Controls.Add(edtTxtScholarship);
        y += 40;

        var lblDur = new Label() { Text = "Training Duration", Top = y, Left = 20, Width = 200, Font = new Font("Segoe UI", 10) };
        edtTxtTrainingDuration = new TextBox() { Top = y, Left = 230, Width = 430, Height = 28, Text = txtTrainingDuration.Text };
        Controls.Add(lblDur);
        Controls.Add(edtTxtTrainingDuration);
        y += 40;

        var lblDate = new Label() { Text = "Date of Assessment", Top = y, Left = 20, Width = 200, Font = new Font("Segoe UI", 10) };
        edtDtAssessment = new DateTimePicker() { Top = y, Left = 230, Width = 200, Value = dtAssessment.Value };
        Controls.Add(lblDate);
        Controls.Add(edtDtAssessment);
        y += 50;

        y += 20;

        btnEditCandidates = new Button()
        {
            Text = "✎ Edit Candidate Details",
            Top = y,
            Left = 20,
            Width = 320,
            Height = 40,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            BackColor = Color.FromArgb(0, 102, 204),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            FlatAppearance = { BorderSize = 0 },
            Cursor = Cursors.Hand
        };
        btnEditCandidates.Click += (s, e) =>
        {
            var candForm = new CandidateForm((int)edtNumCandidates.Value, customCandidates, edtCboQualification.SelectedItem?.ToString() ?? "");
            if (candForm.ShowDialog(this) == DialogResult.OK)
            {
                customCandidates = candForm.GetCandidates();
            }
        };
        Controls.Add(btnEditCandidates);

        btnEditDates = new Button()
        {
            Text = "⚙ Edit Page Dates",
            Top = y,
            Left = 360,
            Width = 300,
            Height = 40,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            BackColor = Color.FromArgb(0, 102, 204),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            FlatAppearance = { BorderSize = 0 },
            Cursor = Cursors.Hand,
            Visible = (int)edtNumCandidates.Value > 10
        };
        edtNumCandidates.ValueChanged += (s, e) => btnEditDates.Visible = (int)edtNumCandidates.Value > 10;
        btnEditDates.Click += (s, e) =>
        {
            int dateCount = (int)Math.Ceiling((int)edtNumCandidates.Value / 10.0) - 1;
            var datesForm = new CustomDatesForm(dateCount, edtDtAssessment.Value, customDates);
            if (datesForm.ShowDialog(this) == DialogResult.OK)
            {
                customDates = datesForm.GetDates();
            }
        };
        Controls.Add(btnEditDates);
        y += 60;

        btnEditSignatories = new Button()
        {
            Text = "👤 Edit Signatories",
            Top = y,
            Left = 20,
            Width = 300,
            Height = 40,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            BackColor = Color.FromArgb(76, 175, 80),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            FlatAppearance = { BorderSize = 0 },
            Cursor = Cursors.Hand
        };
        btnEditSignatories.Click += (s, e) =>
        {
            var sigForm = new SignatoryEditorForm(signatories);
            if (sigForm.ShowDialog(this) == DialogResult.OK)
            {
                signatories = sigForm.GetSignatories();
            }
        };
        Controls.Add(btnEditSignatories);
        y += 60;

        y += 20;

        btnDone = new Button()
        {
            Text = "✓ Done - Regenerate",
            Top = y,
            Left = 150,
            Width = 140,
            Height = 40,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            BackColor = Color.FromArgb(40, 167, 69),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            FlatAppearance = { BorderSize = 0 },
            Cursor = Cursors.Hand
        };
        btnDone.Click += (s, e) =>
        {
            if (!ValidateEditFormFields())
            {
                return;
            }

            numCandidates.Value = (int)edtNumCandidates.Value;
            cboQualification.SelectedIndex = edtCboQualification.SelectedIndex;
            txtAssessor.Text = edtTxtAssessor.Text;
            txtRQMCode.Text = edtTxtRQMCode.Text;
            txtScholarship.Text = edtTxtScholarship.Text;
            txtTrainingDuration.Text = edtTxtTrainingDuration.Text;
            dtAssessment.Value = edtDtAssessment.Value;
            
            DialogResult = DialogResult.OK;
            Close();
        };
        Controls.Add(btnDone);

        btnCancel = new Button()
        {
            Text = "✕ Cancel",
            Top = y,
            Left = 310,
            Width = 140,
            Height = 40,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            BackColor = Color.FromArgb(220, 53, 69),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            FlatAppearance = { BorderSize = 0 },
            Cursor = Cursors.Hand
        };
        btnCancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };
        Controls.Add(btnCancel);
    }

    private bool ValidateEditFormFields()
    {
        List<string> errors = new List<string>();

        int N = (int)edtNumCandidates.Value;
        if (N == 0)
        {
            errors.Add("Number of candidates is required");
        }

        if (edtCboQualification.SelectedIndex < 0)
        {
            errors.Add("Qualification must be selected");
        }

        if (string.IsNullOrWhiteSpace(edtTxtAssessor.Text))
        {
            errors.Add("Assessor Name is required");
        }

        if (string.IsNullOrWhiteSpace(edtTxtRQMCode.Text))
        {
            errors.Add("RQM Code is required");
        }

        if (string.IsNullOrWhiteSpace(edtTxtScholarship.Text))
        {
            errors.Add("Type of Scholarship/Modality is required");
        }

        if (string.IsNullOrWhiteSpace(edtTxtTrainingDuration.Text))
        {
            errors.Add("Training Duration is required");
        }

        if (N > 0)
        {
            if (customCandidates.Count == 0)
            {
                errors.Add("Candidate Details are required - Click 'Edit Candidate Details' to add candidates");
            }
            else if (customCandidates.Count < N)
            {
                errors.Add($"Only {customCandidates.Count} out of {N} candidate(s) entered - please fill in all candidate details");
            }
            else
            {
                bool hasMissingData = false;
                for (int i = 0; i < customCandidates.Count; i++)
                {
                    var candidate = customCandidates[i];
                    if (string.IsNullOrWhiteSpace(candidate.Name) || candidate.Name.Contains("Candidate"))
                    {
                        hasMissingData = true;
                        break;
                    }
                    if (string.IsNullOrWhiteSpace(candidate.Reference) || candidate.Reference.Contains("AC-"))
                    {
                        hasMissingData = true;
                        break;
                    }
                }

                if (hasMissingData)
                {
                    errors.Add("Some candidate names or references are still blank - please fill in all candidate details");
                }
            }
        }

        if (errors.Count > 0)
        {
            string errorMessage = string.Join("\n", errors);
            MessageBox.Show(errorMessage, "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
        }

        return true;
    }

    public List<Candidate> GetCandidates() => customCandidates;
    public List<DateTime> GetDates() => customDates;
    public Signatory GetSignatories() => signatories;
}

// ====================================================================
// SIGNATORY EDITOR FORM (DIALOG)
// ====================================================================
class SignatoryEditorForm : Form
{
    private Signatory signatories;
    private TextBox txtAttendanceSheetName = new TextBox();
    private TextBox txtAttendanceSheetPosition = new TextBox();
    private TextBox txtBillingAssessorACName = new TextBox();
    private TextBox txtBillingAssessorACPosition = new TextBox();
    private TextBox txtBillingAssessorName = new TextBox();
    private TextBox txtBillingAssessorPosition = new TextBox();
    private TextBox txtBillingAssessmentACName = new TextBox();
    private TextBox txtBillingAssessmentACPosition = new TextBox();
    private TextBox txtBillingAssessmentVSName = new TextBox();
    private TextBox txtBillingAssessmentVSPosition = new TextBox();

    public SignatoryEditorForm(Signatory sig)
    {
        signatories = sig ?? new Signatory();
        
        Text = "Edit Signatories";
        Width = 600;
        Height = 850;
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        BackColor = Color.FromArgb(240, 242, 245);
        Font = new Font("Segoe UI", 10);

        int y = 20;

        AddLabel(20, y, "ATTENDANCE SHEET");
        y += 30;
        AddField("Name:", 20, y, txtAttendanceSheetName, signatories.AttendanceSheetName);
        y += 40;
        AddField("Position:", 20, y, txtAttendanceSheetPosition, signatories.AttendanceSheetPosition);
        y += 50;

        AddLabel(20, y, "BILLING ASSESSOR FORM - Approved by:");
        y += 30;
        AddField("Name:", 20, y, txtBillingAssessorACName, signatories.BillingAssessorACName);
        y += 40;
        AddField("Position:", 20, y, txtBillingAssessorACPosition, signatories.BillingAssessorACPosition);
        y += 50;

        AddLabel(20, y, "BILLING ASSESSOR FORM - Prepared by:");
        y += 30;
        AddField("Name:", 20, y, txtBillingAssessorName, signatories.BillingAssessorName);
        y += 40;
        AddField("Position:", 20, y, txtBillingAssessorPosition, signatories.BillingAssessorPosition);
        y += 50;

        AddLabel(20, y, "BILLING ASSESSMENT FORM - Prepared by:");
        y += 30;
        AddField("Name:", 20, y, txtBillingAssessmentACName, signatories.BillingAssessmentACName);
        y += 40;
        AddField("Position:", 20, y, txtBillingAssessmentACPosition, signatories.BillingAssessmentACPosition);
        y += 50;

        AddLabel(20, y, "BILLING ASSESSMENT FORM - Note by:");
        y += 30;
        AddField("Name:", 20, y, txtBillingAssessmentVSName, signatories.BillingAssessmentVSName);
        y += 40;
        AddField("Position:", 20, y, txtBillingAssessmentVSPosition, signatories.BillingAssessmentVSPosition);
        y += 50;

        var btnOK = new Button
        {
            Text = "OK",
            Left = 200,
            Top = y,
            Width = 90,
            Height = 40,
            BackColor = Color.FromArgb(70, 150, 250),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        btnOK.FlatAppearance.BorderSize = 0;
        btnOK.Click += (s, e) =>
        {
            signatories.AttendanceSheetName = txtAttendanceSheetName.Text.Trim().ToUpper();
            signatories.AttendanceSheetPosition = txtAttendanceSheetPosition.Text.Trim();
            signatories.BillingAssessorACName = txtBillingAssessorACName.Text.Trim().ToUpper();
            signatories.BillingAssessorACPosition = txtBillingAssessorACPosition.Text.Trim();
            signatories.BillingAssessorName = txtBillingAssessorName.Text.Trim().ToUpper();
            signatories.BillingAssessorPosition = txtBillingAssessorPosition.Text.Trim();
            signatories.BillingAssessmentACName = txtBillingAssessmentACName.Text.Trim().ToUpper();
            signatories.BillingAssessmentACPosition = txtBillingAssessmentACPosition.Text.Trim();
            signatories.BillingAssessmentVSName = txtBillingAssessmentVSName.Text.Trim().ToUpper();
            signatories.BillingAssessmentVSPosition = txtBillingAssessmentVSPosition.Text.Trim();
            DialogResult = DialogResult.OK;
            Close();
        };
        Controls.Add(btnOK);

        var btnCancel = new Button
        {
            Text = "Cancel",
            Left = 305,
            Top = y,
            Width = 90,
            Height = 40,
            BackColor = Color.FromArgb(200, 200, 200),
            ForeColor = Color.Black,
            Font = new Font("Segoe UI", 10),
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        btnCancel.FlatAppearance.BorderSize = 0;
        btnCancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };
        Controls.Add(btnCancel);
    }

    private void AddLabel(int left, int top, string text)
    {
        var lbl = new Label
        {
            Text = text,
            Left = left,
            Top = top,
            Width = 500,
            Height = 25,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            ForeColor = Color.FromArgb(60, 100, 150)
        };
        Controls.Add(lbl);
    }

    private void AddField(string labelText, int left, int top, TextBox textBox, string value)
    {
        var lbl = new Label
        {
            Text = labelText,
            Left = left,
            Top = top,
            Width = 80,
            Height = 25,
            Font = new Font("Segoe UI", 10)
        };
        Controls.Add(lbl);

        textBox.Left = left + 90;
        textBox.Top = top;
        textBox.Width = 450;
        textBox.Height = 30;
        textBox.Text = value;
        textBox.Font = new Font("Segoe UI", 10);
        Controls.Add(textBox);
    }

    public Signatory GetSignatories() => signatories;
}

// ====================================================================
// GENERATION TITLE FORM (DIALOG)
// ====================================================================
class GenerationTitleForm : Form
{
    private TextBox txtTitle = new TextBox();
    public string FolderTitle { get; private set; } = "";

    public GenerationTitleForm()
    {
        Text = "Enter Folder Title";
        Width = 450;
        Height = 250;
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;

        var lblPrompt = new Label
        {
            Text = "Enter a title or name for this generation:",
            Left = 25,
            Top = 25,
            Width = 400,
            Height = 30,
            Font = new Font("Segoe UI", 10)
        };
        Controls.Add(lblPrompt);

        txtTitle.Left = 25;
        txtTitle.Top = 65;
        txtTitle.Width = 400;
        txtTitle.Height = 35;
        txtTitle.Font = new Font("Segoe UI", 10);
        txtTitle.Focus();
        Controls.Add(txtTitle);

        var btnOK = new Button
        {
            Text = "OK",
            Left = 170,
            Top = 140,
            Width = 100,
            Height = 40,
            BackColor = Color.FromArgb(70, 150, 250),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        btnOK.FlatAppearance.BorderSize = 0;
        btnOK.Click += (s, e) =>
        {
            if (string.IsNullOrWhiteSpace(txtTitle.Text))
            {
                MessageBox.Show("Please enter a title for the generation", "Input Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            FolderTitle = txtTitle.Text.Trim();
            DialogResult = DialogResult.OK;
            Close();
        };
        Controls.Add(btnOK);

        var btnCancel = new Button
        {
            Text = "Cancel",
            Left = 285,
            Top = 140,
            Width = 100,
            Height = 40,
            BackColor = Color.FromArgb(200, 200, 200),
            ForeColor = Color.Black,
            Font = new Font("Segoe UI", 10),
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        btnCancel.FlatAppearance.BorderSize = 0;
        btnCancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };
        Controls.Add(btnCancel);

        KeyPreview = true;
        KeyDown += (s, e) =>
        {
            if (e.KeyCode == Keys.Enter)
            {
                btnOK.PerformClick();
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Escape)
            {
                DialogResult = DialogResult.Cancel;
                Close();
            }
        };
    }
}
