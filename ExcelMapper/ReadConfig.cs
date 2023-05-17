using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using AutoMapper.Execution;
using ExcelMapper.Attributes;
using ExcelMapper.CellOption;
using ExcelMapper.Interface;

namespace ExcelMapper
{
    /// <summary>
    ///     表格读取设置
    /// </summary>
    public sealed class ReadConfig<T> : ExcelConfig<T, ReadCellOption<T>>
    {
        private readonly Stream _excelStream;

        public ReadConfig()
        {
            Init = null;
        }

        /// <summary>
        ///     根据文件路径的初始化
        /// </summary>
        /// <param name="filepath"> 文件路径 </param>
        public ReadConfig(string filepath) : this()
        {
            _excelStream = new MemoryStream();
            using var fs = new FileStream(filepath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite);
            fs.CopyTo(_excelStream);
        }

        /// <summary>
        ///     根据文件流的初始化
        /// </summary>
        /// <param name="stream">文件流</param>
        public ReadConfig(Stream stream) : this()
        {
            _excelStream = new MemoryStream();
            stream.CopyTo(_excelStream);
        }

        /// <summary>
        ///     读取成功之后调用的针对T的委托
        /// </summary>
        public Func<T, T> Init { get; private set; }

        /// <summary>
        ///     添加默认单元格读取设置(其实就是不读取Excel直接给T的某个字段赋值)
        /// </summary>
        /// <param name="prop">T的属性</param>
        /// <param name="defaultValue">默认值</param>
        public ReadConfig<T> Default<TE>(Expression<Func<T, TE>> prop, TE defaultValue)
        {
            Add(GenOption(string.Empty, prop, _ => defaultValue));
            return this;
        }

        /// <summary>
        ///     check条件为True时添加默认单元格读取设置(其实就是不读取Excel直接给T的某个字段赋值)
        /// </summary>
        /// <param name="check">判断结果</param>
        /// <param name="prop">T的属性</param>
        /// <param name="defaultValue">默认值</param>
        public ReadConfig<T> DefaultIf<TE>(bool check, Expression<Func<T, TE>> prop, TE defaultValue)
        {
            return check ? Default(prop, defaultValue) : this;
        }

        /// <summary>
        ///     读取设置处理
        /// </summary>
        /// <param name="field">表头列</param>
        /// <param name="prop">T的属性</param>
        /// <param name="action">对单元格字符串的操作</param>
        public ReadConfig<T> Handler<TE>(string field, Expression<Func<T, TE>> prop, Func<string, TE>? action = null)
        {
            ReadCellOption<T>? option = null;
            var fieldOptions = new List<ReadCellOption<T>>();
            if (FieldOption.Any(t => t.ExcelField == field))
            {
                FieldOption.ToList().ForEach(t =>
                {
                    if (t.ExcelField == field)
                    {
                        t = GenOption(field, prop, action);
                        option = t;
                        fieldOptions.Add(option);
                    }
                    else
                    {
                        fieldOptions.Add(t);
                    }
                });
            }


            if (option == null)
            {
                option = GenOption(field, prop, action);
                Add(option);
            }
            else
            {
                FieldOption = fieldOptions;
            }

            return this;
        }

        /// <summary>
        ///     check条件为True时 读取设置处理
        /// </summary>
        /// <param name="check">判断结果</param>
        /// <param name="field">表头列</param>
        /// <param name="prop">T的属性</param>
        /// <param name="action">对单元格字符串的操作</param>
        public ReadConfig<T> Handler<TE>(bool check, string field, Expression<Func<T, TE>> prop,
            Func<string, TE>? action = null)
        {
            return check ? Handler(field, prop, action) : this;
        }

        /// <summary>
        ///     添加普通单元格设置
        /// </summary>
        /// <param name="field">表头列</param>
        /// <param name="prop">T的属性</param>
        /// <param name="action">对单元格字符串的操作</param>
        public ReadConfig<T> Add<TE>(string field, Expression<Func<T, TE>> prop, Func<string, TE>? action = null)
        {
            Add(GenOption(field, prop, action));
            return this;
        }

        /// <summary>
        ///     check条件为True时添加单元格设置
        /// </summary>
        /// <param name="check">判断结果</param>
        /// <param name="field">表头列</param>
        /// <param name="prop">T的属性</param>
        /// <param name="action">对单元格字符串的操作</param>
        public ReadConfig<T> AddIf<TE>(bool check, string field, Expression<Func<T, TE>> prop,
            Func<string, TE>? action = null)
        {
            return check ? Add(field, prop, action) : this;
        }

        /// <summary>
        ///     添加普通单元格设置
        /// </summary>
        /// <param name="field">表头列</param>
        /// <param name="prop">属性</param>
        /// <param name="action">对单元格字符串的操作</param>
        public ReadConfig<T> Add(string field, PropertyInfo prop, Func<string, object>? action = null)
        {
            Add(GenOption(field, prop, action));
            return this;
        }

        /// <summary>
        ///     添加普通单元格设置
        /// </summary>
        /// <param name="check">判断结果</param>
        /// <param name="field">表头列</param>
        /// <param name="prop">属性</param>
        /// <param name="action">对单元格字符串的操作</param>
        public ReadConfig<T> AddIf(bool check, string field, PropertyInfo prop, Func<string, object>? action = null)
        {
            return check ? Add(field, prop, action) : this;
        }

        /// <summary>
        ///     添加普通单元格设置
        /// </summary>
        /// <param name="field">表头列</param>
        /// <param name="propName">属性名称</param>
        public ReadConfig<T> Add(string field, string propName)
        {
            return Add(field, typeof(T).GetProperty(propName)!);
        }

        /// <summary>
        ///     添加普通单元格设置
        /// </summary>
        /// <param name="check">判断结果</param>
        /// <param name="field">表头列</param>
        /// <param name="propName">属性名称</param>
        public ReadConfig<T> AddIf(bool check, string field, string propName)
        {
            return check ? Add(field, propName) : this;
        }

        /// <summary>
        ///     添加行数据初始化
        /// </summary>
        /// <param name="action"></param>
        public ReadConfig<T> AddInit(Func<T, T> action)
        {
            Init = action;
            return this;
        }

        /// <summary>
        ///     添加行数据初始化
        /// </summary>
        /// <param name="check">判断结果</param>
        /// <param name="action"></param>
        public ReadConfig<T> AddInitIf(bool check, Func<T, T> action)
        {
            return check ? AddInit(action) : this;
        }

        /// <summary>
        ///     生成单元格设置
        /// </summary>
        /// <param name="field"></param>
        /// <param name="prop"></param>
        /// <param name="action"></param>
        public ReadCellOption<T> GenOption<TE>(string field, Expression<Func<T, TE>> prop, Func<string, TE>? action)
        {
            return GenOption(field, (PropertyInfo)prop.GetMember(), action);
        }

        /// <summary>
        ///     生成单元格设置
        /// </summary>
        /// <param name="field"></param>
        /// <param name="prop"></param>
        public ReadCellOption<T> GenOption(string field, PropertyInfo prop)
        {
            return new ReadCellOption<T>
            {
                ExcelField = field,
                Prop = prop
            };
        }

        /// <summary>
        ///     生成单元格设置
        /// </summary>
        /// <param name="field"></param>
        /// <param name="prop"></param>
        /// <param name="action"></param>
        public ReadCellOption<T> GenOption<TE>(string field, PropertyInfo prop, Func<string, TE>? action)
        {
            return action == null
                ? GenOption(field, prop)
                : new ReadCellOption<T>
                {
                    ExcelField = field,
                    Prop = prop,
                    Action = item => action(item)
                };
        }

        /// <summary>
        ///     根据 T 生成默认的 Config
        /// </summary>
        public static ReadConfig<T> GenDefaultConfig()
        {
            // 根据 T 中设置的 ExcelAttribute 创建导入配置
            var value = typeof(T).AttrValues<ExcelAttribute>();
            return value.Any()
                ? GenDefaultConfigByAttribute()
                : GenDefaultConfigByProps();
        }

        /// <summary>
        ///     根据 T 中设置的 ExcelAttribute 创建导入配置
        /// </summary>
        public static ReadConfig<T> GenDefaultConfigByAttribute()
        {
            var config = new ReadConfig<T>();
            var attrData = typeof(T).AttrValues<ExcelAttribute>();
            foreach (var prop in attrData)
            {
                ReadCellOption<T>? option = null;
                var attributes = prop.Key.GetCustomAttributes(typeof(ExcelAttribute), false);
                foreach (var attribute in attributes)
                {
                    if (attribute is ExcelAttribute excelAttribute &&
                        !string.IsNullOrEmpty(excelAttribute.ReadConverterExp)
                        && !string.IsNullOrEmpty(excelAttribute.Separator))
                    {
                        var converterExps = excelAttribute.ReadConverterExp.Split(excelAttribute.Separator);
                        option = config.GenOption(prop.Value.ExcelField, prop.Key,
                            item =>
                            {
                                var value = converterExps.FirstOrDefault(t => t.Split("=")[1] == item);
                                return value == null ? item : value.Split("=")[0];
                            });
                    }
                }

                option ??= config.GenOption(prop.Value.ExcelField, prop.Key);
                config.Add(option);
            }

            return config;
        }

        /// <summary>
        ///     直接根据属性名称创建导入配置
        /// </summary>
        public static ReadConfig<T> GenDefaultConfigByProps()
        {
            var config = new ReadConfig<T>();
            // 直接根据属性名称创建导入配置
            foreach (var prop in typeof(T).GetProperties())
            {
                config.Add(config.GenOption(prop.Name, prop));
            }
            return config;
        }

        #region 异步转换到实体

        /// <summary>
        ///     将表格数据转换为T类型的集合
        /// </summary>
        public async Task<IEnumerable<T>> ToEntityAsync(IExcelReader sheet)
        {
            var header = sheet.HeadersWithIndex;
            var rowCount = sheet.RowCount;
            ConcurrentBag<T> data = new();
            await Task.Factory.StartNew(() =>
            {
                Parallel.For(1, rowCount, index =>
                {
                    Monitor.Enter(sheet);
                    var dataRow = sheet[index].ToList();
                    Monitor.Exit(sheet);
                    // 根据对应传入的设置 为obj赋值
                    if (!dataRow.Any())
                    {
                        return;
                    }

                    var obj = Activator.CreateInstance<T>();
                    foreach (var option in FieldOption)
                    {
                        if (!string.IsNullOrEmpty(option.ExcelField))
                        {
                            if (!header.ContainsKey(option.ExcelField))
                            {
                                continue;
                            }
                            var value = dataRow[header[option.ExcelField]];
                            var targetType = (option.Prop.PropertyType.IsGenericType &&
                                              option.Prop.PropertyType.GetGenericTypeDefinition() ==
                                              typeof(Nullable<>)
                                ? Nullable.GetUnderlyingType(option.Prop.PropertyType)
                                : option.Prop.PropertyType) ?? typeof(string);
                            object propertyVal;
                            if (option.Action != null)
                            {
                                propertyVal = Convert.ChangeType(option.Action.Invoke(value), targetType);
                            }
                            else
                            {
                                propertyVal = value != null ? Convert.ChangeType(value, targetType) : null;
                            }
                            option.Prop.SetValue(obj, propertyVal);
                        }
                        else
                        {
                            option.Prop.SetValue(obj, option.Action?.Invoke(string.Empty));
                        }
                    }

                    Init?.Invoke(obj);
                    data.Add(obj);
                });
            });
            return data;
        }

        /// <summary>
        ///     转换到实体
        /// </summary>
        public async Task<IEnumerable<T>> ToEntityAsync()
        {
            return await ToEntityAsync(_excelStream);
        }

        /// <summary>
        ///     转换到实体
        /// </summary>
        public async Task<IEnumerable<T>> ToEntityAsync(string path)
        {
            await using var fs = new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite);
            return await ToEntityAsync(fs);
        }

        /// <summary>
        ///     转换到实体
        /// </summary>
        public async Task<IEnumerable<T>> ToEntityAsync(Stream stream)
        {
            using var reader = ExcelType.ReaderFunc(stream);
            return await ToEntityAsync(reader);
        }

        /// <summary>
        ///     转换到实体
        /// </summary>
        public static async Task<IEnumerable<T>> ExcelToEntityAsync(string path)
        {
            await using var fs = new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite);
            return await ExcelToEntityAsync(fs);
        }

        /// <summary>
        ///     转换到实体
        /// </summary>
        public static async Task<IEnumerable<T>> ExcelToEntityAsync(Stream stream)
        {
            var reader = ExcelType.ReaderFunc(stream);
            return await ExcelToEntityAsync(reader);
        }

        /// <summary>
        ///     转换到实体
        /// </summary>
        public static async Task<IEnumerable<T>> ExcelToEntityAsync(IExcelReader reader)
        {
            var config = GenDefaultConfig();
            return await config.ToEntityAsync(reader);
        }

        #endregion

        #region 同步转换到实体

        /// <summary>
        ///     将表格数据转换为T类型的集合(更快)
        /// </summary>
        public IEnumerable<T> ToEntity(IExcelReader sheet)
        {
            var header = sheet.HeadersWithIndex;
            var rowCount = sheet.RowCount;
            foreach (var index in Enumerable.Range(1, (int)rowCount - 1))
            {
                var dataRow = sheet[index].ToList();
                // 根据对应传入的设置 为obj赋值
                if (!dataRow.Any())
                {
                    continue;
                }

                var obj = Activator.CreateInstance<T>();
                foreach (var option in FieldOption)
                {
                    if (!string.IsNullOrEmpty(option.ExcelField))
                    {
                        if (!header.ContainsKey(option.ExcelField))
                        {
                            continue;
                        }
                        var value = dataRow[header[option.ExcelField]];
                        var targetType = (option.Prop.PropertyType.IsGenericType &&
                                          option.Prop.PropertyType.GetGenericTypeDefinition() ==
                                          typeof(Nullable<>)
                            ? Nullable.GetUnderlyingType(option.Prop.PropertyType)
                            : option.Prop.PropertyType) ?? typeof(string);
                        var propertyVal = Convert.ChangeType(option.Action.Invoke(value), targetType);
                        option.Prop.SetValue(obj, propertyVal);
                    }
                    else
                    {
                        option.Prop.SetValue(obj, option.Action.Invoke(string.Empty));
                    }
                }

                Init?.Invoke(obj);
                yield return obj;
            }
        }

        /// <summary>
        ///     转换到实体(更快)
        /// </summary>
        public IEnumerable<T> ToEntity()
        {
            return ToEntity(_excelStream);
        }

        /// <summary>
        ///     转换到实体(更快)
        /// </summary>
        public IEnumerable<T> ToEntity(string path)
        {
            using var fs = new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite);
            return ToEntity(fs);
        }

        /// <summary>
        ///     转换到实体(更快)
        /// </summary>
        public IEnumerable<T> ToEntity(Stream stream)
        {
            using var reader = ExcelType.ReaderFunc(stream);
            return ToEntity(reader);
        }

        /// <summary>
        ///     转换到实体(更快)
        /// </summary>
        public static IEnumerable<T> ExcelToEntity(string path)
        {
            using var fs = new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite);
            return ExcelToEntity(fs);
        }

        /// <summary>
        ///     转换到实体(更快)
        /// </summary>
        public static IEnumerable<T> ExcelToEntity(Stream stream)
        {
            var reader = ExcelType.ReaderFunc(stream);
            return ExcelToEntity(reader);
        }

        /// <summary>
        ///     转换到实体(更快)
        /// </summary>
        public static IEnumerable<T> ExcelToEntity(IExcelReader reader)
        {
            var config = GenDefaultConfig();
            return config.ToEntity(reader);
        }

        #endregion
    }
}