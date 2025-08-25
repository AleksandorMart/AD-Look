using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using System.Windows.Forms;
using System.Collections;
using Microsoft.VisualBasic.Devices;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;
using Microsoft.VisualBasic.ApplicationServices;
using System.Drawing;

namespace ADlook
{
    public partial class Form1 : Form
    {
        // ������ ��� ������ � Active Directory
        private readonly AdService _adService = new AdService();

        // ������� ��� �������� ���������������� ��������
        private readonly Dictionary<string, PredefinedQuery> _queries = new Dictionary<string, PredefinedQuery>();

        // ������� ��������� ������
        private PredefinedQuery _currentQuery;

        // ������� ��������� � ���������� ���������� ������
        private string[] freeProperties = { "displayName", "sAMAccountName", "pwdLastSet", "distinguishedName" };
        private string[] freeDisplayNames = { "���", "�����", "��������� ����� ������", "������������" };

        public Form1()
        {
            InitializeComponent();

            // ������������� ����������� ����������
            SetupUI();

            // �������� ���������������� ��������
            InitializePredefinedQueries();

            // ��������� ����������� ������
            InitializeComboBox();
        }

        // ��������� ��������� ����������
        private void SetupUI()
        {
            // ��������� DataGridView
            dgvResults.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvResults.ReadOnly = true;
            dgvResults.AllowUserToAddRows = false;
            dgvResults.ShowCellToolTips = true;

            // ��������� ComboBox
            cmbQueries.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbFreeSearch.DropDownStyle = ComboBoxStyle.DropDownList;

            // �������� ������ ���������� ������ �� ���������
            pnlFreeSearch.Visible = false;
            pnlOUSearch.Visible = false;
            pnlTimeSearch.Visible = false;
            pnlOSSearch.Visible = false;
        }

        // ������������� ���������������� ��������
        private void InitializePredefinedQueries()
        {
            _queries.Add("1. ���������� ��", new PredefinedQuery(
                filter: $"(objectCategory=computer)(lastLogon<={DateTime.Now.AddMonths(-3).ToFileTime()})",
                properties: new[] { "displayName", "lastLogon", "operatingSystem", "distinguishedName" },
                displayNames: new[] { "��� ��", "��������� ����", "��", "������������" }
                ));
            _queries.Add("2. ���������� ��", new PredefinedQuery(
                filter: "(objectCategory=computer)(|(operatingSystem=*Windows 7 ����������������*)(operatingSystem=*Windows XP Professional*))",
                properties: new[] { "displayName", "operatingSystem", "operatingSystemVersion", "whenChanged", "distinguishedName" },
                displayNames: new[] { "��� ��", "��", "������ ��", "��������� ����������", "������������" }
                ));
            _queries.Add("3. ���������� ��", new PredefinedQuery(//TimestampTimestamp
                filter: $"(objectCategory=user)(lastLogon<={DateTime.Now.AddMonths(-6).ToFileTime()})(!userAccountControl:1.2.840.113556.1.4.803:=2)",
                properties: new[] { "displayName", "sAMAccountName", "lastLogon", "distinguishedName", "userAccountControl" },
                displayNames: new[] { "���", "�����", "��������� ����", "������������", "����" }
                ));
            _queries.Add("4. �� ��� ��", new PredefinedQuery(
                filter: "(objectCategory=computer)",
                properties: new[] { "displayName", "dNSHostName", "distinguishedName", "memberOf" },
                displayNames: new[] { "��� ��", "�������� ���", "������������", "������" }
                ));
            _queries.Add("5. �� ��� �������", new PredefinedQuery(
                filter: "(objectCategory=user)(userAccountControl:1.2.840.113556.1.4.803:=32)(!userAccountControl:1.2.840.113556.1.4.803:=2)",
                properties: new[] { "displayName", "sAMAccountName", "distinguishedName", "userAccountControl" },
                displayNames: new[] { "���", "�����", "������������", "����" }
                ));
            _queries.Add("6. �� � ������", new PredefinedQuery(
                filter: "(objectCategory=user)(userAccountControl:1.2.840.113556.1.4.803:=2)(msExchModerationFlags>=1)",
                properties: new[] { "displayName", "mail", "msExchModerationFlags", "distinguishedName", "userAccountControl" },
                displayNames: new[] { "���", "�����", "���������", "������������", "����" }
                ));
            _queries.Add("7. � RW � RO", new PredefinedQuery(
                filter: "(objectCategory=user)(!userAccountControl:1.2.840.113556.1.4.803:=2)",
                properties: new[] { "displayName", "memberOf", "distinguishedName", "userAccountControl" },
                displayNames: new[] { "���", "������", "������������", "����" }
                ));
            _queries.Add("8. �� � WDS_Drop", new PredefinedQuery(
                filter: "(objectCategory=computer)",
                properties: new[] { "displayName", "whenCreated", "distinguishedName" },
                displayNames: new[] { "��� ��", "����� �������", "������������" }
                ));
            _queries.Add("9. ������ ������", new PredefinedQuery(
                filter: "(objectCategory=user)(userAccountControl:1.2.840.113556.1.4.803:=65536)(!userAccountControl:1.2.840.113556.1.4.803:=2)",
                properties: new[] { "displayName", "sAMAccountName", "pwdLastSet", "distinguishedName", "userAccountControl" },
                displayNames: new[] { "���", "�����", "��������� ����� ������", "������������", "����" }
                ));
            _queries.Add("10. ��������� �����", new PredefinedQuery(
                filter: "(objectCategory=user)(objectClass=*)",
                properties: freeProperties,
                displayNames: freeDisplayNames
                ));
            _queries.Add("11. ����� ������", new PredefinedQuery(
                filter: "(objectCategory=user)(!userAccountControl:1.2.840.113556.1.4.803:=2)",
                properties: new[] { "displayName", "sAMAccountName", "memberOf", "distinguishedName" },
                displayNames: new[] { "���", "�����", "������", "������������" }
                ));
        }

        // ���������� ����������� ������ ���������
        private void InitializeComboBox()
        {
            cmbQueries.BeginUpdate();
            cmbQueries.Items.Clear();

            foreach (var key in _queries.Keys)
            {
                cmbQueries.Items.Add(key);
            }
            cmbQueries.SelectedIndex = 0;

            cmbFreeSearch.BeginUpdate();
            cmbFreeSearch.Items.Clear();
            cmbFreeSearch.Items.Add("user");
            cmbFreeSearch.Items.Add("computer");
            cmbFreeSearch.SelectedIndex = 0;
        }

        // ���������� ��������� ���������� �������
        private void cmbQueries_SelectedIndexChanged(object sender, EventArgs e)
        {
            // ������� ���������� ����������
            dgvResults.Columns.Clear();
            dgvResults.Rows.Clear();
            lblStatus.Text = "�������: 0 �������";

            string selectedKey = cmbQueries.SelectedItem.ToString();

            _currentQuery = _queries[selectedKey];

            // ���������� ������ ������ ��� ���������� ������
            pnlFreeSearch.Visible = (selectedKey == "10. ��������� �����");
            pnlOUSearch.Visible = (selectedKey == "10. ��������� �����" || selectedKey == "9. ������ ������" ||
                selectedKey == "1. ���������� ��" || selectedKey == "2. ���������� ��" ||
                selectedKey == "3. ���������� ��" || selectedKey == "11. ����� ������");
            pnlTimeSearch.Visible = (selectedKey == "1. ���������� ��" || selectedKey == "3. ���������� ��");
            pnlOSSearch.Visible = (selectedKey == "2. ���������� ��" || selectedKey == "11. ����� ������");

            if (pnlFreeSearch.Visible)
            {
                cmbFreeSearch.SelectedIndex = 0;
            }
            if (pnlOUSearch.Visible)
            {
                txt3.Text = "";
            }
            if (pnlTimeSearch.Visible)
            {
                dtpTimeSearcher.Value = DateTime.Now;
            }
            if (pnlOSSearch.Visible)
            {
                if (selectedKey == "2. ���������� ��")
                {
                    label5.Text = "��";
                    label6.Text = "���������� �����";
                    label7.Text = "�������";
                }
                if (selectedKey == "11. ����� ������")
                {
                    label5.Text = "������";
                    label6.Text = " ";
                    label7.Text = " ";
                }
                txt4.Text = "";
            }
        }

        // ��������� ������� ������
        private void ExecuteQuery()
        {
            var results = _adService.Search(_currentQuery.Filter, _currentQuery.Properties);
            DisplayResults(results);
        }

        // ����������� ����������� � DataGridView
        private void DisplayResults(List<AdObject> results)
        {
            // ������� ���������� ����������
            dgvResults.Columns.Clear();
            dgvResults.Rows.Clear();

            // �������� ��������
            for (int i = 0; i < _currentQuery.DisplayNames.Length; i++)
            {
                dgvResults.Columns.Add(
                    columnName: $"col_{i}",
                    headerText: _currentQuery.DisplayNames[i]
                );
            }

            // ������������� ��������
            // ����� 1
            if (cmbQueries.SelectedIndex == 0)
            {
                FilterOU(results, txt3.Text);
            }

            // ����� 2
            if (cmbQueries.SelectedIndex == 1)
            {
                FilterOU(results, txt3.Text);
            }

            // ����� 3
            if (cmbQueries.SelectedIndex == 2)
            {
                FilterOU(results, txt3.Text);
            }

            // ������ 4
            if (cmbQueries.SelectedIndex == 3)
            {
                FilterSoftwareGroups(results);
            }

            // ������ 7
            if (cmbQueries.SelectedIndex == 6)
            {
                FilterRWRO(results);// �������� �������� ����� �������� �����
            }

            // ������ 8
            if (cmbQueries.SelectedIndex == 7)
            {
                FilterOU(results);
            }

            // ����� 9
            //if (cmbQueries.SelectedIndex == 8)
            //{
            //    //FilterUsersWithMissingOu(results);
            //    FilterOneMoreGroup(results);
            //}

            // ����� 9
            if (cmbQueries.SelectedIndex == 8)
            {
                FilterOU(results, txt3.Text);
            }

            // ����� 10
            if (cmbQueries.SelectedIndex == 9)
            {
                if (!string.IsNullOrWhiteSpace(txt3.Text))
                    FilterOU(results, txt3.Text);
            }

            // ����� 11
            if (cmbQueries.SelectedIndex == 10)
            {
                if (!string.IsNullOrWhiteSpace(txt3.Text))
                    FilterOU(results, txt3.Text);

                if (!string.IsNullOrWhiteSpace(txt4.Text))
                    FilterGroup(results, txt4.Text);
            }

            // ��������� �������
            foreach (var item in results)
            {
                int rowIndex = dgvResults.Rows.Add();
                var row = dgvResults.Rows[rowIndex];

                for (int i = 0; i < _currentQuery.Properties.Length; i++)
                {
                    row.Cells[i].Value = item.GetPropertyValue(_currentQuery.Properties[i]);
                }
            }

            lblStatus.Text = $"�������: {results.Count} �������";
        }

        // ���������� ����������� �� ������� ��������� ��
        private void FilterSoftwareGroups(List<AdObject> computers)
        {
            var requiredFullGroups = new List<string>
            {
                "GPCP_Additional_software_Install",
                "GPCP_Kaspersky_Endpoint_Security_Install",
                //"GPCP_OCS_Agent_Install",
                //"GPCP_3CX_Phone_Client_Install",
                //""
            };

            // ������� ������� ��� �������� ������
            var requiredGroupsSet = new HashSet<string>(
                requiredFullGroups,
                StringComparer.OrdinalIgnoreCase);

            // ������� ���������� ������� ������� �� ���� ������������ �������
            computers.RemoveAll(computer =>
            {
                // �������� ������ ����� ����������
                string groupsValue = computer.GetPropertyValue("memberOf")?.ToString() ?? "";
                if (string.IsNullOrEmpty(groupsValue))
                    return false;

                // ��������� ������� ���� ������������ �����
                foreach (var requiredGroup in requiredGroupsSet)
                {
                    // ���� ������ � ������ ����� ����������
                    bool found = false;
                    foreach (var computerGroup in groupsValue.Split(';'))
                    {
                        if (computerGroup.Trim().Equals(requiredGroup, StringComparison.OrdinalIgnoreCase))
                        {
                            found = true;
                            break;
                        }
                    }

                    // ���� �� ����� ������ ���� ������ ��������� �� � �����������
                    if (!found) return false;
                }

                // ������� ��������� �� �����������, �� ��� ������ �������
                return true;
            });
        }

        // ���������� ����������� � _WDS Drop
        private void FilterOU(List<AdObject> computers, string ou = "_WDS Drop")
        {
            computers.RemoveAll(computer =>
            {
                string dn = computer.GetPropertyValue("distinguishedName").ToString();
                return !dn.Contains($"OU={ou}", StringComparison.OrdinalIgnoreCase);
            });
        }

        // ���������� �� ������� ������
        private void FilterGroup(List<AdObject> results, string group)
        {
            // ������� ������ ������� �� ����� ������ ������
            results.RemoveAll(res =>
            {
                // �������� ������ �����
                string groupsValue = res.GetPropertyValue("memberOf")?.ToString() ?? "";
                if (string.IsNullOrEmpty(groupsValue))
                    return true;

                bool flag = false;

                // ��������� ������� ������
                foreach (var Group in groupsValue.Split(';'))
                {
                    if (!string.IsNullOrEmpty(Group) && Group[0] == ' ')
                    {
                        if (Group.Substring(1).Equals(group, StringComparison.OrdinalIgnoreCase))
                        {
                            flag = true;
                        }
                    }
                    else
                    {
                        if (Group.Equals(group, StringComparison.OrdinalIgnoreCase))
                        {
                            flag = true;
                        }
                    }

                    // ���� ����� ������ ��������� ���������
                    if (flag) return false;
                }

                // ������� ���������, �� �� ����� ������
                return true;
            });
        }

        // ��������� �������������, ������� ���� �� ���� ���� ����� RW/RO
        private void FilterRWRO(List<AdObject> users)
        {
            //������� ������������� ��� ������ �����
            users.RemoveAll(users =>
            {
                // �������� ������ �����
                string groupsValue = users.GetPropertyValue("memberOf").ToString() ?? "";
                var userGroups = groupsValue.Split(";").Select(g => g.Trim()).ToList();

                // �������� ������� ����� �����
                var baseNames = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

                foreach (var userGroup in userGroups)
                {
                    // ��������� �������� RW/ RO
                    if (userGroup.EndsWith("_RW", StringComparison.OrdinalIgnoreCase) ||
                    userGroup.EndsWith("_RO", StringComparison.OrdinalIgnoreCase))
                    {
                        // ��������� ������� ��� ������
                        string baseName = userGroup.Substring(0, userGroup.Length - 3);

                        // �������������� ��������� ��� �������� �����
                        if (!baseNames.ContainsKey(baseName))
                            baseNames[baseName] = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                        // ��������� ������� (RW ��� RO)
                        string suffix = userGroup.Substring(userGroup.Length - 2);
                        baseNames[baseName].Add(suffix);
                    }
                }

                // ��������� ������� ������ �����
                foreach (var suffixes in baseNames.Values)
                {
                    // ���� ��� �������� ����� ���� ��� ��������
                    if (suffixes.Contains("RW") && suffixes.Contains("RO"))
                        return false; // ��������� ������������ � �����������
                }

                // ��� ������ ����� - ������� ������������
                return true;
            });
        }

        // ��������� �������������� �������� GRUS_*, �� �������� ���������������� OU � DN
        private void FilterUsersWithMissingOu(List<AdObject> users)
        {
            users.RemoveAll(user =>
            {
                // ����������� ������ � ���� � ������
                string dn = user.GetPropertyValue("distinguishedName")?.ToString() ?? "";
                string groupsValue = user.GetPropertyValue("memberOf")?.ToString() ?? "";

                // ������� ������ ���������� �����
                var problemGroups = new List<string>();

                // ���������� ������� �������� ������ � ����
                foreach (var group in groupsValue.Split(';'))
                {
                    string trimmed = group.Trim();
                    if (trimmed.StartsWith("GRUS_", StringComparison.OrdinalIgnoreCase))
                    {
                        string groupName = trimmed.Substring(5);
                        string ouPattern = $"OU={groupName},";

                        if (dn.IndexOf(ouPattern, StringComparison.OrdinalIgnoreCase) < 0)
                        {
                            problemGroups.Add(trimmed);
                        }
                    }
                }

                // ���� ���� ���������� ������ - ��������� �� � ��������� ������������
                if (problemGroups.Count > 0)
                {
                    user.SetProperty("ProblemGroups", string.Join("; ", problemGroups));
                    return false;
                }

                return true;
            });
        }

        // ������� �� ������� ����� ��� ����� ������ ������
        private void FilterOneMoreGroup(List<AdObject> users)
        {
            users.RemoveAll(user =>
            {
                // �������� ������ �����
                string groupsValue = user.GetPropertyValue("memberOf")?.ToString() ?? "";

                int count = 0;

                foreach (var group in groupsValue.Split(';'))
                {
                    // ���� ������ ��������� � GRUS_ ����������� �������
                    string Tgroup = group.Trim();
                    if (Tgroup.StartsWith("GRUS_", StringComparison.OrdinalIgnoreCase) && !Tgroup.EndsWith("_HEAD", StringComparison.OrdinalIgnoreCase))
                    {
                        count++;
                    }
                }

                // ���� ������������ ������ � ����� ������ ������, ������� ���
                if (count <= 1)
                {
                    return true;
                }

                // ���� ����� ��� � ����� ��� �� � ����� ���������
                return false;
            });
        }

        // ���������� ������ ���������� ���������� ������
        private void btnFreeSearch_Click(object sender, EventArgs e)
        {
            if (pnlFreeSearch.Visible || pnlTimeSearch.Visible || (pnlOSSearch.Visible && label5.Text == "��"))
            {
                // ���������� ������� �� ������ ��������� ������
                _currentQuery.Filter = BuildFreeSearchFilter();
            }

            // ���������� �������
            ExecuteQuery();
        }

        // ������ ������ ��� ���������� ������
        private string BuildFreeSearchFilter()
        {
            if (pnlFreeSearch.Visible)
            {
                var filters = new List<string>();

                // ������ �� ������
                if (!string.IsNullOrWhiteSpace(txt1.Text))
                    filters.Add($"({freeProperties[0]}=*{txt1.Text}*)");

                // ������ �� �����
                if (!string.IsNullOrWhiteSpace(txt2.Text))
                    filters.Add($"({freeProperties[1]}=*{txt2.Text}*)");

                // ������� ������ ��� �������������
                /* const */
                string baseFilter = $"(objectCategory={cmbFreeSearch.SelectedItem.ToString()})";

                // �������������� �������
                if (filters.Count == 0)
                    return $"{baseFilter}(objectClass=*)";

                return $"(&{baseFilter}{string.Join("", filters)})";
            }

            if (pnlTimeSearch.Visible)
            {
                string filter = "";
                string selectedKey = cmbQueries.SelectedItem.ToString();
                long fileTime = dtpTimeSearcher.Value.ToFileTime();

                if (selectedKey == "1. ���������� ��")
                {
                    filter = $"(objectCategory=computer)(lastLogon<={fileTime})";
                }
                if (selectedKey == "3. ���������� ��")
                {
                    filter = $"(objectCategory=user)(lastLogon<={fileTime})(!userAccountControl:1.2.840.113556.1.4.803:=2)";
                }

                return filter;
            }

            if (pnlOSSearch.Visible)
            {
                var filters = new List<string>();

                foreach (var os in txt4.Text.Split(','))
                {
                    if (!string.IsNullOrEmpty(os) && os[0] == ' ')
                    {
                        filters.Add($"(operatingSystem=*{os.Substring(1)}*)");
                    }
                    else
                    {
                        filters.Add($"(operatingSystem=*{os}*)");
                    }
                }

                return $"(objectCategory=computer)(|{string.Join("", filters)})";
            }

            return "";
        }

        // ���������� ��������� ���������� ���������� �������
        private void cmbFreeSearch_SelectedIndexChanged(object sender, EventArgs e)
        {
            bool selectedKey = cmbQueries.SelectedIndex == 9;
            string selectedCategory = cmbFreeSearch.SelectedItem.ToString();

            txt1.Text = "";
            txt2.Text = "";
            txt3.Text = "";

            if (selectedCategory == "user" && selectedKey)
            {
                label1.Text = "�����";
                label2.Text = "���";
                freeProperties = new[] { "displayName", "sAMAccountName", "pwdLastSet", "distinguishedName" };
                freeDisplayNames = new[] { "���", "�����", "��������� ����� ������", "������������" };
                _currentQuery.Properties = freeProperties;
                _currentQuery.DisplayNames = freeDisplayNames;
            }
            if (selectedCategory == "computer" && selectedKey)
            {
                label2.Text = "�������� ��";
                label1.Text = "��";
                freeProperties = new[] { "displayName", "operatingSystem", "distinguishedName" };
                freeDisplayNames = new[] { "��� ��", "��", "������������" };
                _currentQuery.Properties = freeProperties;
                _currentQuery.DisplayNames = freeDisplayNames;
            }

        }

        // ���������� �������� ������� � �������
        private void dgvResults_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;

            var cell = dgvResults.Rows[e.RowIndex].Cells[e.ColumnIndex];
            string fullText = cell.Value?.ToString() ?? "";

            // ������� ����� ��� �����������
            using (Form textViewer = new Form())
            {
                textViewer.Text = "Full text";
                textViewer.Size = new Size(500, 300);

                System.Windows.Forms.TextBox tb = new System.Windows.Forms.TextBox
                {
                    Multiline = true,
                    Dock = DockStyle.Fill,
                    ScrollBars = System.Windows.Forms.ScrollBars.Both,
                    Text = fullText,
                    ReadOnly = true,
                    Font = new System.Drawing.Font("Consolas", 10)
                };

                textViewer.Controls.Add(tb);
                textViewer.ShowDialog();
            }
        }

        // ���������� ������ ��������
        private void btnExport_Click(object sender, EventArgs e)
        {
            using (var dialog = new SaveFileDialog())
            {
                dialog.Filter = "Excel Files|*.xlsx";
                dialog.FileName = $"AD_Request_{cmbQueries.SelectedIndex + 1}_{DateTime.Now:dd.MM.yyyy}.xlsx";

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    ExportToExcel(dialog.FileName);
                }
            }
        }

        // ������� ������ � Excel
        private void ExportToExcel(string filename)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            ProgressForm progressForm = null;

            try
            {
                // ������������� ����� ���������
                progressForm = new ProgressForm("������� � Excel", "���������� ������...");
                progressForm.Show();
                System.Windows.Forms.Application.DoEvents();

                // ������� ��������� Excel
                excelApp = new Microsoft.Office.Interop.Excel.Application();
                excelApp.DisplayAlerts = false;
                excelApp.Visible = false;
                excelApp.ScreenUpdating = false;

                // ������� ����� �����
                progressForm.UpdateMessage("�������� ��������� Excel...");
                workbook = excelApp.Workbooks.Add(Type.Missing);

                // �������� ������ ����
                worksheet = (Worksheet)workbook.Sheets[1];

                // ������������� ��� �����
                string sheetName = cmbQueries.SelectedItem?.ToString() ?? "AD_Data";
                if (sheetName.Length > 31) sheetName = sheetName.Substring(0, 31);
                worksheet.Name = sheetName;

                // ��������� ���������
                for (int col = 0; col < dgvResults.Columns.Count; col++)
                {
                    if (progressForm.Cancelled) break;

                    worksheet.Cells[1, col + 1] = dgvResults.Columns[col].HeaderText;

                    // ����������� ���������
                    Microsoft.Office.Interop.Excel.Range headerCell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[1, col + 1];
                    headerCell.Font.Bold = true;
                    headerCell.Borders.Weight = XlBorderWeight.xlThick;
                }

                // ��������� ������
                int totalRows = dgvResults.Rows.Count;
                int currentRow = 0;
                int rowIndex = 2;
                progressForm.UpdateMessage("������� ������...");
                foreach (DataGridViewRow dgvRow in dgvResults.Rows)
                {
                    if (progressForm.Cancelled) break;
                    if (dgvRow.IsNewRow) continue;

                    currentRow++;

                    // ��������� �������� ������ 10 ����� ��� ��� ������ ������ ���� ����� ������ 50 ��� ���� �������� ������ 10 �����
                    if (currentRow % 10 == 0 || totalRows < 50 || (totalRows - currentRow) <= 10)
                    {
                        progressForm.UpdateProgress(currentRow, totalRows);
                    }

                    for (int col = 0; col < dgvResults.Columns.Count; col++)
                    {
                        // �������� �������� ������
                        object value = dgvRow.Cells[col].Value;

                        // ��������� ��� ������ ��� ����������� ��������������
                        if (value is DateTime dateValue)
                        {
                            worksheet.Cells[rowIndex, col + 1] = dateValue;
                            Microsoft.Office.Interop.Excel.Range dateCell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[rowIndex, col + 1];
                            dateCell.NumberFormat = "dd.mm.yyyy hh:mm";
                        }
                        else
                        {
                            worksheet.Cells[rowIndex, col + 1] = value?.ToString();
                        }
                    }
                    rowIndex++;
                }

                if (progressForm.Cancelled)
                {
                    progressForm.UpdateMessage("������ ��������...");
                    return;
                }

                // ���� ��������� ������ ��������
                progressForm.UpdateMessage("����������� ��������...");
                worksheet.Columns.AutoFit();

                // ��������� ������
                progressForm.UpdateMessage("���������� ��������...");
                Microsoft.Office.Interop.Excel.Range usedRange = worksheet.UsedRange;
                usedRange.AutoFilter(1, Type.Missing, XlAutoFilterOperator.xlAnd, Type.Missing, true);

                // ��������� ����
                progressForm.UpdateMessage("���������� �����...");
                workbook.SaveAs(filename, XlFileFormat.xlOpenXMLWorkbook);

                progressForm.UpdateMessage("������� ��������!");
                progressForm.UpdateProgress(currentRow, totalRows);

                MessageBox.Show($"������ ������� �������������� �:\n{filename}",
                    "������� ��������",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"������ ��� �������� � Excel:\n{ex.Message}",
                    "������",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            finally
            {
                // ������� ��������
                if (workbook != null && !progressForm.Cancelled)
                {
                    workbook.Close(false);
                    Marshal.ReleaseComObject(workbook);
                }

                if (excelApp != null)
                {
                    excelApp.ScreenUpdating = true;
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }

                // ����������� COM �������
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);

                // ��������� ����� ���������
                if (progressForm != null)
                {
                    if (!progressForm.IsDisposed)
                    {
                        progressForm.Close();
                        progressForm.Dispose();
                    }
                }

                // �������������� ������ ������ ��� COM ��������
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }

    internal class PredefinedQuery
    {
        public string Filter { get; set; }
        public string[] Properties { get; set; }
        public string[] DisplayNames { get; set; }

        public PredefinedQuery(string filter, string[] properties, string[] displayNames)
        {
            Filter = filter;
            Properties = properties;
            DisplayNames = displayNames;
        }
    }
}