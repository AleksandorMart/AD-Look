using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using System.Xml.Linq;
using System.Collections;
using System.IO;
using System.Windows.Forms;

namespace ADlook
{
    // Сервис для выполнения запросов к Active Directory
    internal class AdService
    {
        // Домен Active Directory (настраивается в конфигурации)
        private readonly string _domain = "LDAP://your.domain";

        // Выполняет поиск в Active Directory по указанному фильтру
        public List<AdObject> Search(string filter, string[] propertiesToLoad)
        {
            var results = new List<AdObject>();

            // Создаем подключение к домену
            using (var entry = new DirectoryEntry(_domain))
            // Создаем поисковый объект
            using (var searcher = new DirectorySearcher(entry))
            {
                // Установка фильтра поиска
                searcher.Filter = $"(&{filter})";
                // Настройка пейджинга для больших результатов
                searcher.PageSize = 1000;

                // Добавляем запрошенные атрибуты
                searcher.PropertiesToLoad.AddRange(propertiesToLoad);

                // Выполняем поиск и обрабатываем результаты
                foreach (SearchResult result in searcher.FindAll())
                {
                    var adObject = new AdObject();
                    // Обрабатываем каждый запрошенный атрибут
                    foreach (var propName in propertiesToLoad)
                    {
                        if (result.Properties.Contains(propName))
                        {
                            var f = result.Properties[propName].Count;
                            object value = result.Properties[propName][0];
                            
                            if (f > 1) 
                            {
                                List<object> res = new List<object>();
                                for (int i = 0; i < f; i++)
                                {
                                    res.Add(result.Properties[propName][i]);
                                }
                                adObject.SetProperty(propName, ConvertAdValue(res, propName));
                            }
                            else
                            {
                                adObject.SetProperty(propName, ConvertAdValue(value, propName));
                            }
                        }
                    }

                    results.Add(adObject);
                }
            }

            return results;
        }

        // Преобразует значение атрибута AD в нужный формат
        private object ConvertAdValue(object value, string propertyName)
        {
            // Обработка memberOf: извлекаем только имена групп
            if (propertyName == "memberOf")
            {
                if (value is IEnumerable values)
                {
                    var groupNames = new List<string>();
                    foreach (var val in values)
                    {
                        string dn = val.ToString();
                        groupNames.Add(ExtractGroupName(dn));//
                    }
                    return string.Join("; ", groupNames);
                }
                return null;
            }

            // Обработка временных меток (FILETIME)
            if (propertyName.EndsWith("Timestamp") ||
                propertyName == "pwdLastSet" ||
                propertyName == "lastLogon" )
            {
                long fileTime = Convert.ToInt64(value);

                // Значение 0 означает "никогда"
                return fileTime > 0 ? DateTime.FromFileTime(fileTime) : (DateTime?)null;
            }

            // Обработка флагов userAccountControl
            if (propertyName == "userAccountControl")
            {
                int flags = Convert.ToInt32(value);
                //return $"0x{flags:X8}"; // Форматирование в HEX
                return flags;
            }

            // Для всех остальных атрибутов - строковое представление
            return value.ToString();
        }

        // Извлечение имени группы из DN
        private string ExtractGroupName(string dn)
        {
            int start = dn.IndexOf("CN=", StringComparison.OrdinalIgnoreCase);
            if (start < 0) return dn;

            start += 3; // Пропускаем "CN="
            int end = dn.IndexOf(",", start);//????
            return end > 0
                ? dn.Substring(start, end - start)
                : dn.Substring(start);
        }
    }

    // Представляет объект Active Directory
    internal class AdObject
    {
        // Словарь для хранения свойств объекта
        private readonly Dictionary<string, object> _properties = new Dictionary<string, object>();

        // Устанавливает значение свойства
        public void SetProperty(string name, object value)
        {
            _properties[name] = value;
        }

        // Возвращает значение свойства
        public object GetPropertyValue(string propertyName)
        {
            return _properties.TryGetValue(propertyName, out object value)
                ? value
                : "N/A";
        }
    }
}

