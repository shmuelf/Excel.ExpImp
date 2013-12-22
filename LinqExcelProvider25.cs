using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Reflection;
//using Castle.DynamicProxy;
using System.ComponentModel;
using System.IO;
namespace ExcelExpImp
{
    public class ExpImp<T> : IExpImp<T>, IDisposable /*where T : INotifyPropertyChanged*/
    {
        public ExpImp(string folderPath)
        {
            TEMP_EXCELS_DIR = folderPath;
        }
        private string TEMP_EXCELS_DIR; //System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData, Environment.SpecialFolderOption.DoNotVerify) + @"\Interfaces\Excel"
        private string fileName;
        public void Export(IEnumerable<T> objs, string name)
        {
            string file = GetTempFilePath(name: name);
            ExcelProvider provider = ExcelProvider.Create(file);
            ExcelSheet<T> sheet = provider.GetSheet<T>();
            foreach (T t in objs)
                sheet.InsertOnSubmit(t);
            provider.SubmitChanges();
            //MemoryStream ms = new MemoryStream();
            //using (var strm = File.OpenRead(file))
            //   strm.CopyTo(ms);
            //File.Delete(file);
            //ms.Position = 0;
            //return ms;
        }

        public Stream Export(IEnumerable<T> objs)
        {
            string file = GetTempFilePath();
            ExcelProvider provider = ExcelProvider.Create(file);
            ExcelSheet<T> sheet = provider.GetSheet<T>();
            foreach (T t in objs)
                sheet.InsertOnSubmit(t);
            provider.SubmitChanges();
            MemoryStream ms = new MemoryStream();
            using (var strm = File.OpenRead(file))
                strm.CopyTo(ms);
            File.Delete(file);
            ms.Position = 0;
            return ms;
        }

        public IEnumerable<T> Import(Stream excel, string excelFileExtension)
        {
            string fileName = GetTempFilePath(excelFileExtension);
            Stream file = null;
            try
            {
                file = File.OpenWrite(fileName); //, FileMode.Create, FileAccess.Write, FileShare.Read
                excel.CopyTo(file);
            }
            catch (Exception ex)
            {
                throw new Exception("Failed to read Excel stream for Import.", ex);
            }
            finally
            {
                if (file!=null) 
                    file.Close();
                excel.Close();
            }

            var res = Import(fileName);
            this.fileName = fileName;
            return res;
        }

        public void Dispose()
        {
            if (File.Exists(fileName))
                File.Delete(fileName);
        }

        public IEnumerable<T> Import(string file)
        {
            ExcelProvider provider = ExcelProvider.Create(file);
            return provider.GetSheet<T>();
        }

        private string GetTempFilePath(string extension = "xlsx", string name = null)
        {
            if(!Directory.Exists(TEMP_EXCELS_DIR))
            {
                var parts = TEMP_EXCELS_DIR.Split('\\');
                var path = "";
                foreach (var dir in parts)
                {
                    path += dir + '\\';
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                }
            }
            name = name ?? Guid.NewGuid().ToString();
            string file;
            while (File.Exists(file = TEMP_EXCELS_DIR + @"\" + name + "." + extension))
            {}
            return file;
        }
    }
    /*[ExcelSheet(Name = "Sheet1")]
    public class Person : INotifyPropertyChanged
    {
        private double id;
        private string fName;
        private string lName;
        private DateTime bDate;
        public event PropertyChangedEventHandler PropertyChanged;
        public Person()
        {
            id = 0;
        }
        protected virtual void SendPropertyChanged(string propertyName)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(propertyName));
            }
        }
        [ExcelColumn(Name = "ID", Storage = "id")]
        public double Id
        {
            get { return id; }
            set { id = value; }
        }
        [ExcelColumn(Name = "First Name", Storage = "fName")]
        public string FirstName
        {
            get { return this.fName; }
            set
            {
                fName = value;
                SendPropertyChanged("FirstName");
            }
        }
        [ExcelColumn(Name = "Last Name", Storage = "lName")]
        public string LastName
        {
            get { return this.lName; }
            set
            {
                lName = value;
                SendPropertyChanged("LastName");
            }
        }
        [ExcelColumn(Name = "BirthDate", Storage = "bDate")]
        public DateTime BirthDate
        {
            get { return this.bDate; }
            set
            {
                bDate = value;
                SendPropertyChanged("BirthDate");
            }
        }
    }*/

    public class ExcelIgnoreColumnAttribute : ExcelColumnAttribute
    {

    }

    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class ExcelColumnAttribute : Attribute
    {
        private string name = string.Empty;
        private string storage = string.Empty;
        private System.Data.OleDb.OleDbType dBType;

        private PropertyInfo propInfo;
        public ExcelColumnAttribute()
        {
            name = string.Empty;
            storage = string.Empty;
            dBType = System.Data.OleDb.OleDbType.Empty;
        }
        
        public string Name
        {
            get { return name; }
            set { name = value; }
        }
        public string Storage
        {
            get { return storage; }
            set { storage = value; }
        }
        public System.Data.OleDb.OleDbType DBType
        {
            get { return dBType; }
            set { dBType = value; }
        }
        internal PropertyInfo GetProperty()
        {
            return propInfo;
        }
        internal void SetProperty(PropertyInfo propInfo)
        {
            this.propInfo = propInfo;
            if (string.IsNullOrEmpty(name))
                name = propInfo.Name;
            if (string.IsNullOrEmpty(storage))
                storage = propInfo.Name;
        }
        internal string GetSelectColumn()
        {
            if (Name == string.Empty)
            {
                return propInfo.Name;
            }
            return Name;
        }
        internal string GetStorageName()
        {
            if (Storage == string.Empty)
            {
                return propInfo.Name;
            }
            return storage;
        }
        internal bool IsFieldStorage()
        {
            return string.IsNullOrEmpty(storage) == false;
        }
    }
    internal class ExcelSheetAttribute : Attribute
    {
        private string name;
        public ExcelSheetAttribute()
        {
        }
        public string Name
        {
            get { return name; }
            set { name = value; }
        }
    }
    internal class ExcelMapReader
    {
        public static string GetSheetName(Type t, OleDbConnection conn = null)
        {
            object[] attr = t.GetCustomAttributes(typeof(ExcelSheetAttribute), true);
            ExcelSheetAttribute sheet;
            if (attr.Length == 0) 
            {
                //throw new InvalidOperationException("ExcelSheetAttribute not found on type " + t.FullName);
                string sheetname = null;
                if (conn != null)
                {
                    //IEnumerable<string> sheets;
                    /*if ((sheets = GetWorksheetNames(conn)).Count()>0)
                        sheetname = sheets.FirstOrDefault();
                    else */if (IsSheetExists(t.Name, conn/*, sheets*/))
                        sheetname = t.Name;
                    else if (IsSheetExists("Sheet1", conn/*, sheets*/))
                        sheetname = "Sheet1";
                    else if (IsSheetExists("גיליון1", conn/*, sheets*/))
                        sheetname = "גיליון1";

                }
                if (sheetname == null)
                    sheetname="Sheet1";
                sheet = new ExcelSheetAttribute { Name = sheetname };
            }
            else
                sheet = (ExcelSheetAttribute)attr[0];
            if (sheet.Name == string.Empty)
                return t.Name;
            return sheet.Name;
        }
        public static List<ExcelColumnAttribute> GetColumnList(Type t)
        {
            List<ExcelColumnAttribute> lst = new List<ExcelColumnAttribute>();
            foreach (PropertyInfo propInfo in t.GetProperties())
            {
                object[] attr = propInfo.GetCustomAttributes(typeof(ExcelColumnAttribute), true);
                if (attr.Length > 0)
                {
                    if (typeof(ExcelIgnoreColumnAttribute) != attr[0].GetType())
                    {
                        ExcelColumnAttribute col = (ExcelColumnAttribute)attr[0];
                        col.SetProperty(propInfo);
                        lst.Add(col);
                    }
                }
                else if (propInfo.GetAccessors().Any(y => y.IsPublic))
                {
                    ExcelColumnAttribute col = new ExcelColumnAttribute { Name = propInfo.Name };
                    col.SetProperty(propInfo);
                    lst.Add(col);
                }
            }
            return lst;
        }

        internal static bool IsSheetExists(string sheet, OleDbConnection conn/*, IEnumerable<string> sheets = null*/)
        {
            /*if (sheets==null)
                sheets = ExcelMapReader.GetWorksheetNames(conn);
            if (sheets.Count() > 0)
                return sheets.Any(n => n == sheet);
            else
            {*/
                try
                {
                    using (OleDbCommand cmd = new OleDbCommand("select 1 from [" + sheet + "$] where 1=0", conn))
                    {
                        var t = cmd.ExecuteScalar();
                        return true;
                    }
                }
                catch (OleDbException ex)
                {
                    return false;
                }
            //}
        }

        internal static IEnumerable<string> GetWorksheetNames(OleDbConnection conn)
        {
            var worksheetNames = new List<string>();
            var excelTables = conn.GetOleDbSchemaTable(
                OleDbSchemaGuid.Tables,
                new Object[] { null, null, null, "TABLE" });

            worksheetNames.AddRange(
                from System.Data.DataRow row in excelTables.Rows
                where row["TABLE_NAME"].ToString().Contains("$")
                let tableName = row["TABLE_NAME"].ToString()
                    .Replace("$", "")
                    .RegexReplace("(^'|'$)", "")
                    .Replace("''", "'")
                where !tableName.Contains("FilterDatabase") && !tableName.Contains("Print_Area")
                select tableName);

            excelTables.Dispose();
            return worksheetNames;
        }
    }
    internal static class Extensions
    {
        internal static string RegexReplace(this string source, string regex, string replacement)
        {
            return System.Text.RegularExpressions.Regex.Replace(source, regex, replacement);
        }
    }

    internal class ExcelConnectionString
    {
        internal static string GetConnectionString(string pFilePath)
        {
            string strConnectionString = string.Empty;
            string strExcelExt = System.IO.Path.GetExtension(pFilePath);

            if (strExcelExt == ".xls")
                strConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties= ""Excel 8.0;HDR=YES;""";
            //Ahad L. Amdani added support for .xslm for workbooks using macros
            else if (strExcelExt == ".xlsx" || strExcelExt == ".xlsm")
                strConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Xml;HDR=YES""";
            else
                throw new ArgumentOutOfRangeException("Excel file extenstion is not known.");

            return string.Format(strConnectionString, pFilePath);

        }
    }
    internal class ExcelSheet<T> : IEnumerable<T>
    {
        private ExcelProvider provider;
        private List<T> rows;
        internal ExcelSheet(ExcelProvider provider)
        {
            this.provider = provider;
            rows = new List<T>();
        }
        private T CreateInstance()
        {
            return Activator.CreateInstance<T>();
        }
        private void Load()
        {
            string connectionString = ExcelConnectionString.GetConnectionString(provider.Filepath);
            List<ExcelColumnAttribute> columns = ExcelMapReader.GetColumnList(typeof(T));
            rows.Clear();
            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                using (OleDbCommand cmd = provider.BuildSelectCommand(conn, typeof(T)))
                {
                    using (OleDbDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            T item = CreateInstance();
                            List<PropertyManager> pm = new List<PropertyManager>();
                            foreach (ExcelColumnAttribute col in columns)
                            {
                                object val = reader[col.GetSelectColumn()];
                                if (val is DBNull)
                                {
                                    val = null;
                                }
                                if (col.IsFieldStorage())
                                {
                                    FieldInfo fi = typeof(T).GetField(col.GetStorageName(), BindingFlags.GetField | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.SetField);
                                    //TomKom 3/13/2009 add change type conversion.
                                    fi.SetValue(item, Convert.ChangeType(val, fi.FieldType));
                                }
                                else
                                {
                                    //TomKom 3/13/2009 add change type conversion.
                                    typeof(T).GetProperty(col.GetStorageName()).SetValue(item, Convert.ChangeType(val, typeof(T).GetProperty(col.GetStorageName()).PropertyType), null);
                                }
                                pm.Add(new PropertyManager(col.GetProperty().Name, val));
                            }
                            rows.Add(item);
                            AddToTracking(item, pm);
                        }
                    }
                }
            }
        }
        private void AddToTracking(Object obj, List<PropertyManager> props)
        {
            provider.ChangeSet.AddObject(new ObjectState(obj, props));
        }
        public void InsertOnSubmit(T entity)
        {
            //Add to tracking
            provider.ChangeSet.InsertObject(entity);
        }
        public void DeleteOnSubmit(T entity)
        {
            provider.ChangeSet.DeleteObject(entity);
        }
        public IEnumerator<T> GetEnumerator()
        {
            Load();
            return rows.GetEnumerator();
        }
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            Load();
            return rows.GetEnumerator();
        }
    }
    internal class ExcelProvider
    {
        private string filePath;
        private ChangeSet changes;
        public ExcelProvider()
        {
            changes = new ChangeSet();
        }
        internal ChangeSet ChangeSet
        {
            get { return changes; }
        }
        internal string Filepath
        {
            get { return filePath; }
        }
        public static ExcelProvider Create(string filePath)
        {
            ExcelProvider provider = new ExcelProvider();
            provider.filePath = filePath;
            return provider;
        }
        public ExcelSheet<T> GetSheet<T>()
        {
            return new ExcelSheet<T>(this);
        }
        public void SubmitChanges()
        {
            string connectionString = ExcelConnectionString.GetConnectionString(this.Filepath);
            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                //conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] {null, null, null, "TABLE"});
                foreach (ObjectState os in this.ChangeSet.ChangedObjects)
                {
                    using (OleDbCommand cmd = BuildCommand(os, conn))
                    {
                        try
                        {
                            cmd.ExecuteNonQuery();
                        }
                        catch (Exception ex)
                        {
                            throw new Exception("sql command failed. command text: " + cmd.CommandText, ex);
                        }
                    }
                }
            }
        }

        public OleDbCommand BuildSelectCommand(OleDbConnection conn, Type entityType)
        {
            OleDbCommand cmd = conn.CreateCommand();
            cmd.CommandText = BuildSelect(entityType, conn);
            return cmd;
        }

        public OleDbCommand BuildCommand(ObjectState os, OleDbConnection conn)
        {
            OleDbCommand cmd = conn.CreateCommand();
            string sheet = ExcelMapReader.GetSheetName(os.Entity.GetType());
            validateSheetName(sheet, cmd, os.Entity.GetType(), conn);
            if (os.ChangeState == ChangeState.Deleted)
            {
                BuildDeleteClause(cmd, os, sheet);
            }
            else if (os.ChangeState == ChangeState.Updated)
            {
                BuildUpdateClause(cmd, os, sheet);
            }
            else if (os.ChangeState == ChangeState.Inserted)
            {
                BuildInsertClause(cmd, os, sheet);
            }
            return cmd;
        }

        private bool validateSheetName(string sheet, OleDbCommand cmd, Type entityType, OleDbConnection conn)
        {
            if (!ExcelMapReader.IsSheetExists(sheet, conn))
            {
                BuildCreateTableClause(sheet, cmd, entityType);
                try
                {
                    cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    throw new Exception("sql command failed. command text: " + cmd.CommandText, ex);
                }
            }
            return false;
        }

        private string GetColumnType(ExcelColumnAttribute col)
        {
            string type;
            if (col.DBType != OleDbType.Empty)
                type = col.DBType.ToString().ToUpper();
            else
            {
                object val;
                if (col.GetProperty().PropertyType == typeof(string))
                    val = string.Empty;
                else
                    val = Activator.CreateInstance(col.GetProperty().PropertyType);
                OleDbParameter para = new OleDbParameter("@1", val);
                type = para.OleDbType.ToString().ToUpper();
            }
            if (type.IndexOf("CHAR") >= 0 || type.IndexOf("TEXT") >= 0)
                type += "(MAX)";
            return type;
        }

        private static void BuildCreateTableClause(string sheet, OleDbCommand cmd, Type entityType)
        {
            StringBuilder sql = new StringBuilder("create table [" + sheet + "] (");
            foreach (ExcelColumnAttribute col in ExcelMapReader.GetColumnList(entityType))
            {
                sql.AppendFormat("[{0}] VARCHAR(50) null" + ",", col.GetSelectColumn()); //" " + GetColumnType(col) + ","); 
            }
            sql.Replace(',', ')', sql.Length - 1, 1);
            cmd.CommandText = sql.ToString();
        }

        private string BuildSelect(Type t, OleDbConnection conn)
        {
            string sheet = ExcelMapReader.GetSheetName(t, conn);
            StringBuilder builder = new StringBuilder();
            foreach (ExcelColumnAttribute col in ExcelMapReader.GetColumnList(t))
            {
                if (builder.Length > 0)
                {
                    builder.Append(", ");
                }
                builder.AppendFormat("[{0}]", col.GetSelectColumn());
            }
            builder.Append(" FROM [");
            builder.Append(sheet);
            builder.Append("$]");
            return "SELECT " + builder.ToString();
        }

        public void BuildInsertClause(OleDbCommand cmd, ObjectState objState, string sheet)
        {
            StringBuilder builder = new StringBuilder();
            builder.Append("INSERT INTO [");
            builder.Append(sheet);
            builder.Append("$]");
            StringBuilder columns = new StringBuilder();
            StringBuilder values = new StringBuilder();
            foreach (ExcelColumnAttribute col in ExcelMapReader.GetColumnList(objState.Entity.GetType()))
            {
                if (columns.Length > 0)
                {
                    columns.Append(", ");
                    values.Append(", ");
                }
                columns.AppendFormat("[{0}]", col.GetSelectColumn());
                string paraNum = "@x" + cmd.Parameters.Count.ToString();
                values.Append(paraNum);
                object val = col.GetProperty().GetValue(objState.Entity, null) ?? string.Empty;
                OleDbParameter para = new OleDbParameter(paraNum, val);
                if (col.DBType != System.Data.OleDb.OleDbType.Empty)
                    para.OleDbType = col.DBType;
                cmd.Parameters.Add(para);
            }
            cmd.CommandText = builder.ToString() + "(" + columns.ToString() + ") VALUES (" +
            values.ToString() + ")";
        }
        public void BuildUpdateClause(OleDbCommand cmd, ObjectState objState, string sheet)
        {
            StringBuilder builder = new StringBuilder();
            builder.Append("UPDATE [");
            builder.Append(sheet);
            builder.Append("$] SET ");
            StringBuilder changeBuilder = new StringBuilder();
            List<ExcelColumnAttribute> cols = ExcelMapReader.GetColumnList(objState.Entity.GetType());
            List<ExcelColumnAttribute> changedCols =
            (from c in cols
             join p in objState.ChangedProperties on c.GetProperty().Name equals p.PropertyName
             where p.HasChanged == true
             select c).ToList();
            foreach (ExcelColumnAttribute col in changedCols)
            {
                if (changeBuilder.Length > 0)
                {
                    changeBuilder.Append(", ");
                }
                string paraNum = "@x" + cmd.Parameters.Count.ToString();
                changeBuilder.AppendFormat("[{0}]", col.GetSelectColumn());
                changeBuilder.Append(" = ");
                changeBuilder.Append(paraNum);
                object val = col.GetProperty().GetValue(objState.Entity, null);
                OleDbParameter para = new OleDbParameter(paraNum, val);
                cmd.Parameters.Add(para);
            }
            builder.Append(changeBuilder.ToString());
            cmd.CommandText = builder.ToString();
            BuildWhereClause(cmd, objState);
        }
        public void BuildDeleteClause(OleDbCommand cmd, ObjectState objState, string sheet)
        {
            StringBuilder builder = new StringBuilder();
            builder.Append("DELETE FROM [");
            builder.Append(sheet);
            builder.Append("$]");
            cmd.CommandText = builder.ToString();
            BuildWhereClause(cmd, objState);
        }
        public void BuildWhereClause(OleDbCommand cmd, ObjectState objState)
        {
            StringBuilder builder = new StringBuilder();
            List<ExcelColumnAttribute> cols = ExcelMapReader.GetColumnList(objState.Entity.GetType());
            foreach (ExcelColumnAttribute col in cols)
            {

                PropertyManager pm = objState.GetProperty(col.GetProperty().Name);
                if (builder.Length > 0)
                {
                    builder.Append(" and ");
                }

                builder.AppendFormat("[{0}]", col.GetSelectColumn());
                //fix from Andrew 4/2/08 to handle empty cells
                if (pm.OrginalValue == System.DBNull.Value)
                    builder.Append(" IS NULL");
                else
                {
                    builder.Append(" = ");
                    string paraNum = "@x" + cmd.Parameters.Count.ToString();
                    builder.Append(paraNum);
                    OleDbParameter para = new OleDbParameter(paraNum, pm.OrginalValue);
                    cmd.Parameters.Add(para);
                }
            }
            cmd.CommandText = cmd.CommandText + " WHERE " + builder.ToString();
        }
    }
    internal class PropertyManager
    {
        private string propertyName;
        private object orginalValue;
        private bool hasChanged;
        public PropertyManager(string propName, object value)
        {
            propertyName = propName;
            orginalValue = value;
            hasChanged = false;
        }
        public string PropertyName
        {
            get { return propertyName; }
            set { propertyName = value; }
        }
        public object OrginalValue
        {
            get { return orginalValue; }
            set { orginalValue = value; }
        }
        public bool HasChanged
        {
            get { return hasChanged; }
            set { hasChanged = value; }
        }
    }
    internal enum ChangeState
    {
        Retrieved,
        Updated,
        Inserted,
        Deleted
    }
    internal class ObjectState
    {
        private List<PropertyManager> propList;
        private object entity;
        private ChangeState state;
        public ObjectState(object entity, List<PropertyManager> props)
        {
            this.entity = entity;
            this.propList = props;
            state = ChangeState.Retrieved;
            if (entity is INotifyPropertyChanged)
            {
                INotifyPropertyChanged i = (INotifyPropertyChanged)entity;
                i.PropertyChanged += new PropertyChangedEventHandler(i_PropertyChanged);
            }
        }
        public List<PropertyManager> Properties
        {
            get { return this.propList; }
        }
        public PropertyManager GetProperty(string propertyName)
        {
            return (from p in propList where p.PropertyName == propertyName select p).FirstOrDefault();
        }
        public List<PropertyManager> ChangedProperties
        {
            get { return (from p in propList where p.HasChanged == true select p).ToList(); }
        }
        public ChangeState ChangeState
        {
            get { return state; }
            set { state = value; }
        }
        public Object Entity
        {
            get { return this.entity; }
        }
        public void i_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            PropertyManager pm = (from p in propList where p.HasChanged == false && p.PropertyName == e.PropertyName select p).FirstOrDefault();
            if (pm != null)
            {
                pm.HasChanged = true;
                if (state == ChangeState.Retrieved)
                    state = ChangeState.Updated;
            }
        }
    }
    internal class ChangeSet
    {
        private List<ObjectState> trackedList;
        public ChangeSet()
        {
            trackedList = new List<ObjectState>();
        }
        public void AddObject(ObjectState objectState)
        {
            trackedList.Add(objectState);
        }
        public void InsertObject(Object obj)
        {
            foreach (ObjectState os in trackedList)
            {
                if (ObjectState.ReferenceEquals(os.Entity, obj))
                {
                    throw new InvalidOperationException("Object already in list");
                }
            }
            ObjectState osNew = new ObjectState(obj, new List<PropertyManager>());
            osNew.ChangeState = ChangeState.Inserted;
            trackedList.Add(osNew);
        }
        public void DeleteObject(Object obj)
        {
            ObjectState os = (from o in trackedList where Object.ReferenceEquals(o.Entity, obj) == true select o).FirstOrDefault();
            if (os != null)
            {
                if (os.ChangeState == ChangeState.Inserted)
                {
                    trackedList.Remove(os);
                }
                else
                {
                    os.ChangeState = ChangeState.Deleted;
                }
            }
        }
        public List<ObjectState> ChangedObjects
        {
            get { return (from c in trackedList where c.ChangeState != ChangeState.Retrieved select c).ToList(); }
        }
    }

    /*public class CastleDM<T>
    {
        public T CreateProxy(Type parent)
        {
            ProxyGenerator generator = new ProxyGenerator();
            return (T)generator.CreateClassProxy(parent, new Type[] { typeof(INotifyPropertyChanged) }, ProxyGenerationOptions.Default, new Object[] { }, new INotifyPropertyChangedInterceptor[] { new INotifyPropertyChangedInterceptor() });
        }
    }

    public class INotifyPropertyChangedInterceptor : StandardInterceptor
    {
        private PropertyChangedEventHandler handler;
        public override object Intercept(IInvocation invocation, params object[] args)
        {
            if (invocation.Method.Name == "add_PropertyChanged")
            {
                return handler = (PropertyChangedEventHandler)Delegate.Combine(handler, (Delegate)args[0]);
            }
            else if (invocation.Method.Name == "remove_PropertyChanged")
            {
                return handler = (PropertyChangedEventHandler)Delegate.Remove(handler, (Delegate)args[0]);
            }
            else if (invocation.Method.Name.StartsWith("set_"))
            {
                if (handler != null) handler(invocation.Proxy, new PropertyChangedEventArgs(invocation.Method.Name.Substring("set_".Length)));
            }
            return base.Intercept(invocation, args);
        }
    }*/

    /*public class DependsOnAttribute : Attribute
    {
        public string[] Properties { get; set; }

        public DependsOnAttribute(params string[] properties)
        {
            this.Properties = properties;
        }
    }*/

    //internal class PropertyChangedInterceptor : StandardInterceptor
    //{
    //    private event PropertyChangedEventHandler _propertyChanged = delegate { };
    //    private bool _proceed;

    //    protected override void PreProceed(IInvocation invocation)
    //    {
    //        _proceed = false;

    //        if (invocation.Method.Name == "add_PropertyChanged")
    //        {
    //            _propertyChanged += invocation.Arguments[0] as PropertyChangedEventHandler;
    //        }
    //        else if (invocation.Method.Name == "remove_PropertyChanged")
    //        {
    //            _propertyChanged -= invocation.Arguments[0] as PropertyChangedEventHandler;
    //        }
    //        else
    //        {
    //            _proceed = true;
    //        }
    //    }
    //    protected override void PerformProceed(IInvocation invocation)
    //    {
    //        if (_proceed)
    //        {
    //            base.PerformProceed(invocation);
    //        }
    //    }
    //    protected override void PostProceed(IInvocation invocation)
    //    {
    //        if (invocation.Method.Name.StartsWith("set_"))
    //        {
    //            _propertyChanged(invocation.Proxy, new PropertyChangedEventArgs(invocation.Method.Name.Substring("set_".Length)));
                /*

                var dependencies = invocation.Proxy.GetType().GetProperties()
                    .Where(p => p.GetCustomAttributes(typeof(DependsOnAttribute), true)
                        .Any(a => ((DependsOnAttribute)a).Properties.Contains(invocation.Method.Name.Substring("set_".Length))));
                foreach (var dependentProperty in dependencies)
                {
                    _propertyChanged(invocation.Proxy, new PropertyChangedEventArgs(dependentProperty.Name));
                }*/
    //        }
    //        base.PostProceed(invocation);
    //    }
    //}
    //public static class Notifiable
    //{
    //    private static readonly ProxyGenerator _generator = new ProxyGenerator();

    //    public static T CreateProxy<T>(T model)
    //    {
    //        var interceptor = new PropertyChangedInterceptor();
    //        var interfaces = new Type[] { typeof(INotifyPropertyChanged) };

    //        return (T)_generator.CreateClassProxyWithTarget(model.GetType(), interfaces, model, interceptor);
    //    }
    //}


}
