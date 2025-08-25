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
        // Сервис для работы с Active Directory
        private readonly AdService _adService = new AdService();

        // Словарь для хранения предопределенных запросов
        private readonly Dictionary<string, PredefinedQuery> _queries = new Dictionary<string, PredefinedQuery>();

        // Текущий выбранный запрос
        private PredefinedQuery _currentQuery;

        // Массивы атрибутов и заголовков свободного поиска
        private string[] freeProperties = { "displayName", "sAMAccountName", "pwdLastSet", "distinguishedName" };
        private string[] freeDisplayNames = { "Имя", "Логин", "Последняя смена пароля", "Расположение" };

        public Form1()
        {
            InitializeComponent();

            // Инициализация компонентов интерфейса
            SetupUI();

            // Загрузка предопределенных запросов
            InitializePredefinedQueries();

            // Настройка выпадающего списка
            InitializeComboBox();
        }

        // Настройка элементов интерфейса
        private void SetupUI()
        {
            // Настройка DataGridView
            dgvResults.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvResults.ReadOnly = true;
            dgvResults.AllowUserToAddRows = false;
            dgvResults.ShowCellToolTips = true;

            // Настройка ComboBox
            cmbQueries.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbFreeSearch.DropDownStyle = ComboBoxStyle.DropDownList;

            // Скрываем панель свободного поиска по умолчанию
            pnlFreeSearch.Visible = false;
            pnlOUSearch.Visible = false;
            pnlTimeSearch.Visible = false;
            pnlOSSearch.Visible = false;
        }

        // Инициализация предопределенных запросов
        private void InitializePredefinedQueries()
        {
            _queries.Add("1. Неактивные ПК", new PredefinedQuery(
                filter: $"(objectCategory=computer)(lastLogon<={DateTime.Now.AddMonths(-3).ToFileTime()})",
                properties: new[] { "displayName", "lastLogon", "operatingSystem", "distinguishedName" },
                displayNames: new[] { "Имя ПК", "Последний вход", "ОС", "Расположение" }
                ));
            _queries.Add("2. Устаревшие ОС", new PredefinedQuery(
                filter: "(objectCategory=computer)(|(operatingSystem=*Windows 7*)(operatingSystem=*Windows XP Professional*))",
                properties: new[] { "displayName", "operatingSystem", "operatingSystemVersion", "whenChanged", "distinguishedName" },
                displayNames: new[] { "Имя ПК", "ОС", "Версия ОС", "Последнее обновление", "Расположение" }
                ));
            _queries.Add("3. Неактивные УЗ", new PredefinedQuery(//TimestampTimestamp
                filter: $"(objectCategory=user)(lastLogon<={DateTime.Now.AddMonths(-6).ToFileTime()})(!userAccountControl:1.2.840.113556.1.4.803:=2)",
                properties: new[] { "displayName", "sAMAccountName", "lastLogon", "distinguishedName", "userAccountControl" },
                displayNames: new[] { "Имя", "Логин", "Последний вход", "Расположение", "Флаг" }
                ));
            _queries.Add("4. ПК без ПО", new PredefinedQuery(
                filter: "(objectCategory=computer)",
                properties: new[] { "displayName", "dNSHostName", "distinguishedName", "memberOf" },
                displayNames: new[] { "Имя ПК", "Доменное имя", "Расположение", "Группы" }
                ));
            _queries.Add("5. УЗ без паролей", new PredefinedQuery(
                filter: "(objectCategory=user)(userAccountControl:1.2.840.113556.1.4.803:=32)(!userAccountControl:1.2.840.113556.1.4.803:=2)",
                properties: new[] { "displayName", "sAMAccountName", "distinguishedName", "userAccountControl" },
                displayNames: new[] { "Имя", "Логин", "Расположение", "Флаг" }
                ));
            _queries.Add("6. УЗ с почтой", new PredefinedQuery(
                filter: "(objectCategory=user)(userAccountControl:1.2.840.113556.1.4.803:=2)(msExchModerationFlags>=1)",
                properties: new[] { "displayName", "mail", "msExchModerationFlags", "distinguishedName", "userAccountControl" },
                displayNames: new[] { "Имя", "Почта", "Модерация", "Расположение", "Флаг" }
                ));
            _queries.Add("7. И RW и RO", new PredefinedQuery(
                filter: "(objectCategory=user)(!userAccountControl:1.2.840.113556.1.4.803:=2)",
                properties: new[] { "displayName", "memberOf", "distinguishedName", "userAccountControl" },
                displayNames: new[] { "Имя", "Группы", "Расположение", "Флаг" }
                ));
            _queries.Add("8. ПК в WDS_Drop", new PredefinedQuery(
                filter: "(objectCategory=computer)",
                properties: new[] { "displayName", "whenCreated", "distinguishedName" },
                displayNames: new[] { "Имя ПК", "Когда создали", "Расположение" }
                ));
            _queries.Add("9. Вечные пароли", new PredefinedQuery(
                filter: "(objectCategory=user)(userAccountControl:1.2.840.113556.1.4.803:=65536)(!userAccountControl:1.2.840.113556.1.4.803:=2)",
                properties: new[] { "displayName", "sAMAccountName", "pwdLastSet", "distinguishedName", "userAccountControl" },
                displayNames: new[] { "Имя", "Логин", "Последняя смена пароля", "Расположение", "Флаг" }
                ));
            _queries.Add("10. Свободный поиск", new PredefinedQuery(
                filter: "(objectCategory=user)(objectClass=*)",
                properties: freeProperties,
                displayNames: freeDisplayNames
                ));
            _queries.Add("11. Юзеры группы", new PredefinedQuery(
                filter: "(objectCategory=user)(!userAccountControl:1.2.840.113556.1.4.803:=2)",
                properties: new[] { "displayName", "sAMAccountName", "memberOf", "distinguishedName" },
                displayNames: new[] { "Имя", "Логин", "Группы", "Расположение" }
                ));
        }

        // Заполнение выпадающего списка запросами
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

        // заполнение выпадающего списка запросами
        private void cmbQueries_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Очищаем предыдущие результаты
            dgvResults.Columns.Clear();
            dgvResults.Rows.Clear();
            lblStatus.Text = "Найдено: 0 записей";

            string selectedKey = cmbQueries.SelectedItem.ToString();

            _currentQuery = _queries[selectedKey];

            // Показываем панель только для свободного поиска
            pnlFreeSearch.Visible = (selectedKey == "10. Свободный поиск");
            pnlOUSearch.Visible = (selectedKey == "10. Свободный поиск" || selectedKey == "9. Вечные пароли" ||
                selectedKey == "1. Неактивные ПК" || selectedKey == "2. Устаревшие ОС" ||
                selectedKey == "3. Неактивные УЗ" || selectedKey == "11. Юзеры группы");
            pnlTimeSearch.Visible = (selectedKey == "1. Неактивные ПК" || selectedKey == "3. Неактивные УЗ");
            pnlOSSearch.Visible = (selectedKey == "2. Устаревшие ОС" || selectedKey == "11. Юзеры группы");

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
                if (selectedKey == "2. Устаревшие ОС")
                {
                    label5.Text = "ОС";
                    label6.Text = "перечислять через";
                    label7.Text = "запятую";
                }
                if (selectedKey == "11. Юзеры группы")
                {
                    label5.Text = "Группа";
                    label6.Text = " ";
                    label7.Text = " ";
                }
                txt4.Text = "";
            }
        }

        // Выполняет текущий запрос
        private void ExecuteQuery()
        {
            var results = _adService.Search(_currentQuery.Filter, _currentQuery.Properties);
            DisplayResults(results);
        }

        // Отображение результатов в DataGridView
        private void DisplayResults(List<AdObject> results)
        {
            // Очищаем предыдущие результаты
            dgvResults.Columns.Clear();
            dgvResults.Rows.Clear();

            // Создание столбцов
            for (int i = 0; i < _currentQuery.DisplayNames.Length; i++)
            {
                dgvResults.Columns.Add(
                    columnName: $"col_{i}",
                    headerText: _currentQuery.DisplayNames[i]
                );
            }

            // Постобработка запросов
            // Запрос 1
            if (cmbQueries.SelectedIndex == 0)
            {
                FilterOU(results, txt3.Text);
            }

            // Запрос 2
            if (cmbQueries.SelectedIndex == 1)
            {
                FilterOU(results, txt3.Text);
            }

            // Запрос 3
            if (cmbQueries.SelectedIndex == 2)
            {
                FilterOU(results, txt3.Text);
            }

            // Запрос 4
            if (cmbQueries.SelectedIndex == 3)
            {
                FilterSoftwareGroups(results);
            }

            // Запрос 7
            if (cmbQueries.SelectedIndex == 6)
            {
                FilterRWRO(results);
            }

            // Запрос 8
            if (cmbQueries.SelectedIndex == 7)
            {
                FilterOU(results);
            }

            // Запрос 9
            //if (cmbQueries.SelectedIndex == 8)
            //{
            //    //FilterUsersWithMissingOu(results);
            //    FilterOneMoreGroup(results);
            //}

            // Запрос 9
            if (cmbQueries.SelectedIndex == 8)
            {
                FilterOU(results, txt3.Text);
            }

            // Запрос 10
            if (cmbQueries.SelectedIndex == 9)
            {
                if (!string.IsNullOrWhiteSpace(txt3.Text))
                    FilterOU(results, txt3.Text);
            }

            // Запрос 11
            if (cmbQueries.SelectedIndex == 10)
            {
                if (!string.IsNullOrWhiteSpace(txt3.Text))
                    FilterOU(results, txt3.Text);

                if (!string.IsNullOrWhiteSpace(txt4.Text))
                    FilterGroup(results, txt4.Text);
            }

            // Заполняем данными
            foreach (var item in results)
            {
                int rowIndex = dgvResults.Rows.Add();
                var row = dgvResults.Rows[rowIndex];

                for (int i = 0; i < _currentQuery.Properties.Length; i++)
                {
                    row.Cells[i].Value = item.GetPropertyValue(_currentQuery.Properties[i]);
                }
            }

            lblStatus.Text = $"Найдено: {results.Count} записей";
        }

        // Фильтрация компьютеров по группам установки ОС
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

            // Создаем словарь для быстрого поиска
            var requiredGroupsSet = new HashSet<string>(
                requiredFullGroups,
                StringComparer.OrdinalIgnoreCase);

            // Удаляем компьютеры, которые состоят во всех обязательных группах
            computers.RemoveAll(computer =>
            {
                // Получаем список групп компьютера
                string groupsValue = computer.GetPropertyValue("memberOf")?.ToString() ?? "";
                if (string.IsNullOrEmpty(groupsValue))
                    return false;

                // Проверяем наличие всех обязательных групп
                foreach (var requiredGroup in requiredGroupsSet)
                {
                    // Ищем группу в списке групп компьютера
                    bool found = false;
                    foreach (var computerGroup in groupsValue.Split(';'))
                    {
                        if (computerGroup.Trim().Equals(requiredGroup, StringComparison.OrdinalIgnoreCase))
                        {
                            found = true;
                            break;
                        }
                    }

                    // Если не нашли хотя бы одну группу — оставляем ПК в результатах
                    if (!found) return false;
                }

                // Удаляем компьютер из результатов, так как все группы найдены
                return true;
            });
        }

        // Фильтрация компьютеров в _WDS Drop
        private void FilterOU(List<AdObject> computers, string ou = "_WDS Drop")
        {
            computers.RemoveAll(computer =>
            {
                string dn = computer.GetPropertyValue("distinguishedName").ToString();
                return !dn.Contains($"OU={ou}", StringComparison.OrdinalIgnoreCase);
            });
        }

        // Фильтрация по наличию отдела
        private void FilterGroup(List<AdObject> results, string group)
        {
            // Удаляем записи, которые не имеют нужной группы
            results.RemoveAll(res =>
            {
                // Получаем список групп
                string groupsValue = res.GetPropertyValue("memberOf")?.ToString() ?? "";
                if (string.IsNullOrEmpty(groupsValue))
                    return true;

                bool flag = false;

                // Проверяем наличие группы
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

                    // Если нашли группу оставляем результат
                    if (flag) return false;
                }

                // Удаляем результат, так как не нашли группу
                return true;
            });
        }

        // Ôèëüòðóåò ïîëüçîâàòåëåé, èìåþùèõ õîòÿ áû îäíó ïàðó ãðóïï RW/RO
        private void FilterRWRO(List<AdObject> users)
        {
            //Óäàëÿåì ïîëüçîâàòåëåé áåç ïàðíûõ ãðóïï
            users.RemoveAll(users =>
            {
                // Ïîëó÷àåì ñïèñîê ãðóïï
                string groupsValue = users.GetPropertyValue("memberOf").ToString() ?? "";
                var userGroups = groupsValue.Split(";").Select(g => g.Trim()).ToList();

                // Ñîáèðàåì áàçîâûå èìåíà ãðóïï
                var baseNames = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

                foreach (var userGroup in userGroups)
                {
                    // Ïðîâåðÿåì ñóôôèêñû RW/ RO
                    if (userGroup.EndsWith("_RW", StringComparison.OrdinalIgnoreCase) ||
                    userGroup.EndsWith("_RO", StringComparison.OrdinalIgnoreCase))
                    {
                        // Èçâëåêàåì áàçîâîå èìÿ ãðóïïû
                        string baseName = userGroup.Substring(0, userGroup.Length - 3);

                        // Èíèöèàëèçèðóåì êîëëåêöèþ äëÿ áàçîâîãî èìåíè
                        if (!baseNames.ContainsKey(baseName))
                            baseNames[baseName] = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                        // Äîáàâëÿåì ñóôôèêñ (RW èëè RO)
                        string suffix = userGroup.Substring(userGroup.Length - 2);
                        baseNames[baseName].Add(suffix);
                    }
                }

                // Ïðîâåðÿåì íàëè÷èå ïàðíûõ ãðóïï
                foreach (var suffixes in baseNames.Values)
                {
                    // Åñëè äëÿ áàçîâîãî èìåíè åñòü îáà ñóôôèêñà
                    if (suffixes.Contains("RW") && suffixes.Contains("RO"))
                        return false; // Îñòàâëÿåì ïîëüçîâàòåëÿ â ðåçóëüòàòàõ
                }

                // Íåò ïàðíûõ ãðóïï - óäàëÿåì ïîëüçîâàòåëÿ
                return true;
            });
        }

        // Ôèëüòðóåò ïîëüçîâàòåëåéñ ãðóïïàìè GRUS_*, íå èìåþùèìè ñîîòâåòñòâóþùåãî OU â DN
        private void FilterUsersWithMissingOu(List<AdObject> users)
        {
            users.RemoveAll(user =>
            {
                // Ïðåîáðàçóåì ãðóïïû è ïóòü â ñòðîêè
                string dn = user.GetPropertyValue("distinguishedName")?.ToString() ?? "";
                string groupsValue = user.GetPropertyValue("memberOf")?.ToString() ?? "";

                // Ñîçäàåì ñïèñîê ïðîáëåìíûõ ãðóïï
                var problemGroups = new List<string>();

                // Ñðàâíèâàåì íàëè÷èå íàçâàíèÿ ãðóïïû â ïóòè
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

                // Åñëè åñòü ïðîáëåìíûå ãðóïïû - ñîõðàíÿåì èõ è îñòàâëÿåì ïîëüçîâàòåëÿ
                if (problemGroups.Count > 0)
                {
                    user.SetProperty("ProblemGroups", string.Join("; ", problemGroups));
                    return false;
                }

                return true;
            });
        }

        // Ñìîòðèì íà íàëè÷èå áîëåå ÷åì îäíîé ãðóïïû îòäåëà
        private void FilterOneMoreGroup(List<AdObject> users)
        {
            users.RemoveAll(user =>
            {
                // Ïîëó÷àåì ñòðîêó ãðóïï
                string groupsValue = user.GetPropertyValue("memberOf")?.ToString() ?? "";

                int count = 0;

                foreach (var group in groupsValue.Split(';'))
                {
                    // Åñëè ãðóïïà íà÷èíåòñÿ ñ GRUS_ óâåëè÷èâàåì ñ÷åò÷èê
                    string Tgroup = group.Trim();
                    if (Tgroup.StartsWith("GRUS_", StringComparison.OrdinalIgnoreCase) && !Tgroup.EndsWith("_HEAD", StringComparison.OrdinalIgnoreCase))
                    {
                        count++;
                    }
                }

                // Åñëè ïîëüçîâàòåëü òîëüêî â îäíîé ãðóïïå îòäåëà, óäàëÿåì åãî
                if (count <= 1)
                {
                    return true;
                }

                // Åñëè áîëåå ÷åì â îäíîé èëè íå â îäíîé îñòàâëÿåì
                return false;
            });
        }

        // Îáðàáîò÷èê êíîïêè âûïîëíåíèÿ ñâîáîäíîãî ïîèñêà
        private void btnFreeSearch_Click(object sender, EventArgs e)
        {
            if (pnlFreeSearch.Visible || pnlTimeSearch.Visible || (pnlOSSearch.Visible && label5.Text == "ÎÑ"))
            {
                // Ïîñòðîåíèå ôèëüòðà íà îñíîâå ââåäåííûõ äàííûõ
                _currentQuery.Filter = BuildFreeSearchFilter();
            }

            // Âûïîëíåíèå çàïðîñà
            ExecuteQuery();
        }

        // Ñòðîèò ôèëüòð äëÿ ñâîáîäíîãî ïîèñêà
        private string BuildFreeSearchFilter()
        {
            if (pnlFreeSearch.Visible)
            {
                var filters = new List<string>();

                // Ôèëüòð ïî ëîãèíó
                if (!string.IsNullOrWhiteSpace(txt1.Text))
                    filters.Add($"({freeProperties[0]}=*{txt1.Text}*)");

                // Ôèëüòð ïî èìåíè
                if (!string.IsNullOrWhiteSpace(txt2.Text))
                    filters.Add($"({freeProperties[1]}=*{txt2.Text}*)");

                // Áàçîâûé ôèëüòð äëÿ ïîëüçîâàòåëåé
                /* const */
                string baseFilter = $"(objectCategory={cmbFreeSearch.SelectedItem.ToString()})";

                // Êîìáèíèðîâàíèå óñëîâèé
                if (filters.Count == 0)
                    return $"{baseFilter}(objectClass=*)";

                return $"(&{baseFilter}{string.Join("", filters)})";
            }

            if (pnlTimeSearch.Visible)
            {
                string filter = "";
                string selectedKey = cmbQueries.SelectedItem.ToString();
                long fileTime = dtpTimeSearcher.Value.ToFileTime();

                if (selectedKey == "1. Íåàêòèâíûå ÏÊ")
                {
                    filter = $"(objectCategory=computer)(lastLogon<={fileTime})";
                }
                if (selectedKey == "3. Íåàêòèâíûå ÓÇ")
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

        // Îáðàáîò÷èê èçìåíåíèÿ âûáðàííîãî ñâîáîäíîãî çàïðîñà
        private void cmbFreeSearch_SelectedIndexChanged(object sender, EventArgs e)
        {
            bool selectedKey = cmbQueries.SelectedIndex == 9;
            string selectedCategory = cmbFreeSearch.SelectedItem.ToString();

            txt1.Text = "";
            txt2.Text = "";
            txt3.Text = "";

            if (selectedCategory == "user" && selectedKey)
            {
                label1.Text = "Ëîãèí";
                label2.Text = "ÔÈÎ";
                freeProperties = new[] { "displayName", "sAMAccountName", "pwdLastSet", "distinguishedName" };
                freeDisplayNames = new[] { "Èìÿ", "Ëîãèí", "Ïîñëåäíÿÿ ñìåíà ïàðîëÿ", "Ðàñïîëîæåíèå" };
                _currentQuery.Properties = freeProperties;
                _currentQuery.DisplayNames = freeDisplayNames;
            }
            if (selectedCategory == "computer" && selectedKey)
            {
                label2.Text = "Íàçâàíèå ÏÊ";
                label1.Text = "ÎÑ";
                freeProperties = new[] { "displayName", "operatingSystem", "distinguishedName" };
                freeDisplayNames = new[] { "Èìÿ ÏÊ", "ÎÑ", "Ðàñïîëîæåíèå" };
                _currentQuery.Properties = freeProperties;
                _currentQuery.DisplayNames = freeDisplayNames;
            }

        }

        // Îáðàáîò÷èê äâîéíîãî íàæàòèÿ â òàáëèöå
        private void dgvResults_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;

            var cell = dgvResults.Rows[e.RowIndex].Cells[e.ColumnIndex];
            string fullText = cell.Value?.ToString() ?? "";

            // Ñîçäàåì ôîðìó äëÿ îòîáðàæåíèÿ
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

        // Îáðàáîò÷èê êíîïêè ýêñïîðòà
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

        // Ýêñïîðò äàííûõ â Excel
        private void ExportToExcel(string filename)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            ProgressForm progressForm = null;

            try
            {
                // Èíèöèàëèçàöèÿ ôîðìû ïðîãðåññà
                progressForm = new ProgressForm("Ýêñïîðò â Excel", "Ïîäãîòîâêà äàííûõ...");
                progressForm.Show();
                System.Windows.Forms.Application.DoEvents();

                // Ñîçäàåì ýêçåìïëÿð Excel
                excelApp = new Microsoft.Office.Interop.Excel.Application();
                excelApp.DisplayAlerts = false;
                excelApp.Visible = false;
                excelApp.ScreenUpdating = false;

                // Ñîçäàåì íîâóþ êíèãó
                progressForm.UpdateMessage("Ñîçäàíèå äîêóìåíòà Excel...");
                workbook = excelApp.Workbooks.Add(Type.Missing);

                // Ïîëó÷àåì ïåðâûé ëèñò
                worksheet = (Worksheet)workbook.Sheets[1];

                // Óñòàíàâëèâàåì èìÿ ëèñòà
                string sheetName = cmbQueries.SelectedItem?.ToString() ?? "AD_Data";
                if (sheetName.Length > 31) sheetName = sheetName.Substring(0, 31);
                worksheet.Name = sheetName;

                // Äîáàâëÿåì çàãîëîâêè
                for (int col = 0; col < dgvResults.Columns.Count; col++)
                {
                    if (progressForm.Cancelled) break;

                    worksheet.Cells[1, col + 1] = dgvResults.Columns[col].HeaderText;

                    // Ôîðìàòèðóåì çàãîëîâêè
                    Microsoft.Office.Interop.Excel.Range headerCell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[1, col + 1];
                    headerCell.Font.Bold = true;
                    headerCell.Borders.Weight = XlBorderWeight.xlThick;
                }

                // Äîáàâëÿåì äàííûå
                int totalRows = dgvResults.Rows.Count;
                int currentRow = 0;
                int rowIndex = 2;
                progressForm.UpdateMessage("Ýêñïîðò äàííûõ...");
                foreach (DataGridViewRow dgvRow in dgvResults.Rows)
                {
                    if (progressForm.Cancelled) break;
                    if (dgvRow.IsNewRow) continue;

                    currentRow++;

                    // Îáíîâëÿåì ïðîãðåññ êàæäûå 10 ñòðîê èëè äëÿ êàæäîé ñòðîêè åñëè ñòðîê ìåíüøå 50 èëè åñëè îñòàëîñü ìåíüøå 10 ñòðîê
                    if (currentRow % 10 == 0 || totalRows < 50 || (totalRows - currentRow) <= 10)
                    {
                        progressForm.UpdateProgress(currentRow, totalRows);
                    }

                    for (int col = 0; col < dgvResults.Columns.Count; col++)
                    {
                        // Ïîëó÷àåì çíà÷åíèå ÿ÷åéêè
                        object value = dgvRow.Cells[col].Value;

                        // Ïðîâåðÿåì òèï äàííûõ äëÿ ïðàâèëüíîãî ôîðìàòèðîâàíèÿ
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
                    progressForm.UpdateMessage("Îòìåíà îïåðàöèè...");
                    return;
                }

                // Àâòî íàñòðîéêà øèðèíû ñòîëáöîâ
                progressForm.UpdateMessage("Îïòèìèçàöèÿ ñòîëáöîâ...");
                worksheet.Columns.AutoFit();

                // Äîáàâëÿåì ôèëüòû
                progressForm.UpdateMessage("Äîáàâëåíèå ôèëüòðîâ...");
                Microsoft.Office.Interop.Excel.Range usedRange = worksheet.UsedRange;
                usedRange.AutoFilter(1, Type.Missing, XlAutoFilterOperator.xlAnd, Type.Missing, true);

                // Ñîõðàíÿåì ôàéë
                progressForm.UpdateMessage("Ñîõðàíåíèå ôàéëà...");
                workbook.SaveAs(filename, XlFileFormat.xlOpenXMLWorkbook);

                progressForm.UpdateMessage("Ýêñïîðò çàâåðøåí!");
                progressForm.UpdateProgress(currentRow, totalRows);

                MessageBox.Show($"Äàííûå óñïåøíî ýêñïîðòèðîâàíû â:\n{filename}",
                    "Ýêñïîðò çàâåðøåí",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Îøèáêà ïðè ýêñïîðòå â Excel:\n{ex.Message}",
                    "Îøèáêà",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            finally
            {
                // Î÷èñòêà ðåñóðñîâ
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

                // Îñâîáîæäàåì COM îáúåêòû
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);

                // Çàêðûâàåì ôîðìó ïðîãðåññà
                if (progressForm != null)
                {
                    if (!progressForm.IsDisposed)
                    {
                        progressForm.Close();
                        progressForm.Dispose();
                    }
                }

                // Ïðèíóäèòåëüíàÿ ñáîðêà ìóñîðà äëÿ COM îáúåêòîâ
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
