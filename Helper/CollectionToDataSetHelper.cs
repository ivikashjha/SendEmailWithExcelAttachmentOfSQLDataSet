namespace SendEmailWithExcelAttachmentOfSQLDataSet
{
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.Reflection;

    public static class CollectionToDataSetHelper
    {
        /// <summary>
        /// This extension function convert collection to data set.
        /// </summary>
        /// <typeparam name="T">This is type of list.</typeparam>
        /// <param name="source">List to be converted to dataset.</param>
        /// <param name="name">dataset name.</param>
        /// <returns>DataSet</returns>
        public static DataSet ConvertToDataSet<T>(this IEnumerable<T> source, string name)
        {
            if (source == null)
                throw new ArgumentNullException("source ");
            if (string.IsNullOrEmpty(name))
                throw new ArgumentNullException("name");
            var converted = new DataSet(name);
            converted.Tables.Add(NewTable(name, source));            
            return converted;
        }

        private static DataTable NewTable<T>(string name, IEnumerable<T> list)
        {
            PropertyInfo[] propInfo = typeof(T).GetProperties();
            DataTable table = Table<T>(name, list, propInfo);
            IEnumerator<T> enumerator = list.GetEnumerator();
            while (enumerator.MoveNext())
                table.Rows.Add(CreateRow<T>(table.NewRow(), enumerator.Current, propInfo));
            return table;
        }

        private static DataRow CreateRow<T>(DataRow row, T listItem, PropertyInfo[] pi)
        {
            foreach (PropertyInfo p in pi)
            {
                if (p.PropertyType == typeof(DateTime?) || p.PropertyType == typeof(DateTime))
                {
                    if (p.GetValue(listItem) != null && ((DateTime)p.GetValue(listItem) == DateTime.MinValue || (DateTime)p.GetValue(listItem) == DateTime.MaxValue))
                        row[p.Name.ToString()] = DateTime.UtcNow;
                    else
                        row[p.Name.ToString()] = p.GetValue(listItem) ?? DateTime.UtcNow;
                }
                else
                {
                    row[p.Name.ToString()] = p.GetValue(listItem) ?? DBNull.Value;
                }
            }
            return row;
        }

        private static DataTable Table<T>(string name, IEnumerable<T> list, PropertyInfo[] pi)
        {
            DataTable table = new DataTable(name);
            foreach (PropertyInfo p in pi)
            {
                if (!p.PropertyType.Equals(null))
                {
                    table.Columns.Add(p.Name, Nullable.GetUnderlyingType(p.PropertyType) ?? p.PropertyType);
                    //table.Columns.Add(p.Name, p.PropertyType);
                }
            }
            return table;
        }
    }
}
