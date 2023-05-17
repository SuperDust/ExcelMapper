using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ExcelMapper.Attributes;
using ExcelMapper.CellOption;
using ExcelMapper.Interface;

namespace ExcelMapper
{
    /// <summary>
    ///     导出表格设置
    /// </summary>
    public sealed class ExportConfig<T> : ExcelConfig<T, ExportCellOption<T>>
    {
        public ExportConfig()
        {
        }

        public ExportConfig(IEnumerable<T> data)
        {
            Data = data;
        }


        /// <summary>
        ///     添加普通单元格设置
        /// </summary>
        /// <param name="field">表头列</param>
        /// <param name="propName">属性名称</param>
        /// <param name="width">宽度</param>
        public ExportConfig<T> Add(string field, string propName,int width=0)
        {
            return Add(field, item => item.GetValue(propName),width);
        }

        /// <summary>
        ///     添加普通单元格设置
        /// </summary>
        /// <param name="check">判断结果</param>
        /// <param name="field">表头列</param>
        /// <param name="propName">属性名称</param>
        /// <param name="width">宽度</param>
        public ExportConfig<T> AddIf(bool check, string field, string propName,int width=0)
        {
            return check ? Add(field, propName,width) : this;
        }

        /// <summary>
        ///     添加普通单元格设置
        /// </summary>
        public ExportConfig<T> Add(string field, Func<T, object> action,int width=0)
        {
            Add(new ExportCellOption<T>
            {
                ExcelField = field,
                Action = action,
                Width=width
            });
            return this;
        }

        /// <summary>
        ///     添加普通单元格设置
        /// </summary>
        public ExportConfig<T> AddIf(bool check, string field, Func<T, object> action,int width=0)
        {
            if (check)
            {
                Add(new ExportCellOption<T>
                {
                    ExcelField = field,
                    Action = action,
                    Width=width
                });
            }
            return this;
        }

        /// <summary>
        ///     普通单元格设置 处理
        /// </summary>
        public ExportConfig<T> Handler(string field, Func<T, object> action)
        {
            if (FieldOption.Any(t => t.ExcelField == field))
            {
                var fieldOptions = new List<ExportCellOption<T>>();
                FieldOption.ToList().ForEach(t =>
                {
                    if (t.ExcelField == field)
                    {
                        t = new ExportCellOption<T>
                        {
                            ExcelField = field,
                            Action = action,
                            Width = t.Width,
                        };
                        fieldOptions.Add(t);
                    }
                    else
                    {
                        fieldOptions.Add(t);
                    }
                });
                FieldOption = fieldOptions;
            }
            else
            {
                Add(new ExportCellOption<T>
                {
                    ExcelField = field,
                    Action = action
                });
            }

            return this;
        }

        /// <summary>
        ///     普通单元格设置 处理
        /// </summary>
        public ExportConfig<T> HandlerIf(bool check, string field, Func<T, object> action)
        {
            if (check)
            {
                Handler(field, action);
            }

            return this;
        }

        /// <summary>
        ///     根据 T 生成默认的 Config
        /// </summary>
        public static ExportConfig<T> GenDefaultConfig(IEnumerable<T>? data = null)
        {
            var value = typeof(T).AttrValues<ExcelAttribute>();
            // 根据 T 中设置的 ExcelAttribute 创建导出配置
            return value.Any()
                ? GenDefaultConfigByAttribute(data)
                : GenDefaultConfigByProps(data); // 直接根据属性名称创建导出配置
        }

        /// <summary>
        ///     根据 T 中设置的 ExcelAttribute 创建导出配置
        /// </summary>
        private static ExportConfig<T> GenDefaultConfigByAttribute(IEnumerable<T>? data = null)
        {
            var enumerable = data?.ToList();
            var config = enumerable == null || enumerable.Count == 0
                ? new ExportConfig<T>()
                : new ExportConfig<T>(enumerable);
            var attrData = typeof(T).AttrValues<ExcelAttribute>();
            foreach (var prop in attrData)
            {
                ExportCellOption<T>? option = null;
                var attributes = prop.Key.GetCustomAttributes(typeof(ExcelAttribute), false);
                foreach (var attribute in attributes)
                {
                    if (attribute is not ExcelAttribute excelAttribute)
                    {
                        continue;
                    }
                    if (excelAttribute.Combox != null && excelAttribute.Combox.Any())
                    {
                        config.AddCombox(prop.Value.ExcelField, excelAttribute.Combox);
                    }
                  
                    var converterExps = excelAttribute.ReadConverterExp.Split(excelAttribute.Separator);
                    option = new ExportCellOption<T>
                    {
                        ExcelField = prop.Value.ExcelField,
                        Action = 
                            (string.IsNullOrEmpty(excelAttribute.ReadConverterExp) ||
                             string.IsNullOrEmpty(excelAttribute.Separator))? item =>item.GetValue(prop.Key.Name)
                            : item =>
                        {
                            var propValue = Convert.ToString(item.GetValue(prop.Key.Name));
                            var value = converterExps.FirstOrDefault(t => t.Split("=")[0] == propValue);
                            return value == null ? propValue : value.Split("=")[1];
                        },
                        Width = prop.Value.Width
                    };
                }

                option ??= new ExportCellOption<T>
                {
                    ExcelField = prop.Value.ExcelField,
                    Action = item => item.GetValue(prop.Key.Name),
                    Width = prop.Value.Width
                };
                config.Add(option);
            }

            return config;
        }

        /// <summary>
        ///     直接根据属性名称创建导出配置
        /// </summary>
        private static ExportConfig<T> GenDefaultConfigByProps(IEnumerable<T>? data = null)
        {
            var enumerable = data?.ToList();
            var config = enumerable == null || enumerable.Count == 0
                ? new ExportConfig<T>()
                : new ExportConfig<T>(enumerable);
            var nameProps = typeof(T).GetProperties();
            foreach (var propName in nameProps.Select(item => item.Name))
            {
                config.Add(propName, propName);
            }
            return config;
        }

        #region 异步导出表头

        /// <summary>
        ///     导出表头
        /// </summary>
        public async Task<Stream> ExportHeaderAsync(IExcelCellReader excel,bool autofit = false)
        {
            await Task.Factory.StartNew(() =>
            {
                foreach (var (value, index) in Header.Select((value, index) => (value, index)))
                {
                    excel[0, index].Value = value;
                    var obj = FieldCombox.FirstOrDefault(t => t.Item1 == value);
                    if (obj.Item2 != null && obj.Item2.Any())
                    {
                        excel.SetCombox(index, 2000, obj.Item2);
                    }
                }
            });
            var dictWidth = new Dictionary<int, int>();
            foreach (var (item, index) in FieldOption.Select((item, i) => (item, i)))
            {
                dictWidth.Add(index, item.Width);
            }
            return excel.GetStream(dictWidth,autofit);
        }

        /// <summary>
        ///     导出表头
        /// </summary>
        public async Task<Stream> ExportHeaderAsync(bool autofit = false)
        {
            using var stream = new MemoryStream();
            return await ExportHeaderAsync(stream,autofit);
        }

        /// <summary>
        ///     导出表头
        /// </summary>
        public async Task<Stream> ExportHeaderAsync(string path,bool autofit = false)
        {
            await using var fs = new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite);
            return await ExportHeaderAsync(fs,autofit);
        }

        /// <summary>
        ///     导出表头
        /// </summary>
        public async Task<Stream> ExportHeaderAsync(Stream stream,bool autofit = false)
        {
            using var read = ExcelType.CellReaderFunc(stream);
            var exportStream = await ExportHeaderAsync(read);
            var dictWidth = new Dictionary<int, int>();
            foreach (var (item, index) in FieldOption.Select((item, i) => (item, i)))
            {
                dictWidth.Add(index, item.Width);
            }
            read.Save(dictWidth,autofit);
            return exportStream;
        }

        #endregion

        #region 同步导出表头

        /// <summary>
        ///     导出表头
        /// </summary>
        public Stream ExportHeader(IExcelCellReader excel,bool autofit = false)
        {
            foreach (var (value, index) in Header.Select((value, index) => (value, index)))
            {
                excel[0, index].Value = value;
                var obj = FieldCombox.FirstOrDefault(t => t.Item1 == value);
                if (obj.Item2 != null && obj.Item2.Any())
                {
                    excel.SetCombox(index, 2000, obj.Item2);
                }
            }

            var dictWidth = new Dictionary<int, int>();
            foreach (var (item, index) in FieldOption.Select((item, i) => (item, i)))
            {
                dictWidth.Add(index, item.Width);
            }
            return excel.GetStream(dictWidth,autofit);
        }

        /// <summary>
        ///     导出表头
        /// </summary>
        public Stream ExportHeader(bool autofit = false)
        {
            using var stream = new MemoryStream();
            return ExportHeader(stream,autofit);
        }

        /// <summary>
        ///     导出表头
        /// </summary>
        public Stream ExportHeader(string path,bool autofit = false)
        {
            using var fs = new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite);
            return ExportHeader(fs,autofit);
        }

        /// <summary>
        ///     导出表头
        /// </summary>
        public Stream ExportHeader(Stream stream,bool autofit = false)
        {
            using var read = ExcelType.CellReaderFunc(stream);
            var exportStream = ExportHeader(read);
            var dictWidth = new Dictionary<int, int>();
            foreach (var (item, index) in FieldOption.Select((item, i) => (item, i)))
            {
                dictWidth.Add(index, item.Width);
            }
            read.Save(dictWidth,autofit);
            return exportStream;
        }

        #endregion

        #region 异步导出excel

        /// <summary>
        ///     导出excel
        /// </summary>
        public async Task<Stream> ExportAsync(IExcelCellReader excel, IEnumerable<T>? data = null,bool autofit = false)
        {
            await Task.Factory.StartNew(() =>
            {
                foreach (var (value, index) in Header.Select((value, index) => (value, index)))
                {
                    excel[0, index].Value = value;
                    var obj = FieldCombox.FirstOrDefault(t => t.Item1 == value);
                    if (obj.Item2 != null && obj.Item2.Any())
                    {
                        excel.SetCombox(index, 2000, obj.Item2);
                    }
                }

                var enumerable = data?.ToList();
                if (enumerable != null && enumerable.Any())
                {
                    foreach (var (cellData, rowIndex) in enumerable.Select((value, index) => (value, index + 1)))
                        foreach (var (item, index) in FieldOption.Select((item, i) => (item, i)))
                        {
                            excel[rowIndex, index].Value = item.Action(cellData);
                        }
                }
                else if (Data.Any())
                {
                    foreach (var (cellData, rowIndex) in  Data.Select((value, index) => (value, index + 1)))
                        foreach (var (item, index) in FieldOption.Select((item, i) => (item, i)))
                        {
                            excel[rowIndex, index].Value = item.Action(cellData);
                        }
                }
            });
            var dictWidth = new Dictionary<int, int>();
            foreach (var (item, index) in FieldOption.Select((item, i) => (item, i)))
            {
                dictWidth.Add(index, item.Width);
            }
            return excel.GetStream(dictWidth,autofit);
        }

        /// <summary>
        ///     数据实体导出为Excel
        /// </summary>
        public async Task<Stream> ExportAsync(IEnumerable<T>? data = null,bool autofit = false)
        {
            using var stream = new MemoryStream();
            return await ExportAsync(stream, data,autofit);
        }

        /// <summary>
        ///     数据实体导出为Excel
        /// </summary>
        public async Task<Stream> ExportAsync(string path, IEnumerable<T>? data = null,bool autofit = false)
        {
            await using var fs = new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite);
            return await ExportAsync(fs, data,autofit);
        }

        /// <summary>
        ///     数据实体导出为Excel
        /// </summary>
        public async Task<Stream> ExportAsync(Stream stream, IEnumerable<T>? data = null,bool autofit = false)
        {
            using var read = ExcelType.CellReaderFunc(stream);
            var exportStream = await ExportAsync(read, data);
            var dictWidth = new Dictionary<int, int>();
            foreach (var (item, index) in FieldOption.Select((item, i) => (item, i)))
            {
                dictWidth.Add(index, item.Width);
            }
            read.Save(dictWidth,autofit);
            return exportStream;
        }

        #endregion

        #region 同步导出excel

        /// <summary>
        ///     导出excel
        /// </summary>
        public Stream Export(IExcelCellReader excel, IEnumerable<T>? data = null,bool autofit = false)
        {
            foreach (var (value, index) in Header.Select((value, index) => (value, index)))
            {
                excel[0, index].Value = value;
                var obj = FieldCombox.FirstOrDefault(t => t.Item1 == value);
                if (obj.Item2 != null && obj.Item2.Any())
                {
                    excel.SetCombox(index, 2000, obj.Item2);
                }
            }

            var enumerable = data?.ToList();
            if (enumerable != null && enumerable.Any())
            {
                foreach (var (cellData, rowIndex) in enumerable.Select((value, index) => (value, index + 1)))
                    foreach (var (item, index) in FieldOption.Select((item, i) => (item, i)))
                    {
                        excel[rowIndex, index].Value = item.Action(cellData);
                    }
            }
            else if (Data.Any())
            {
                foreach (var (cellData, rowIndex) in  Data.Select((value, index) => (value, index + 1)))
                    foreach (var (item, index) in FieldOption.Select((item, i) => (item, i)))
                    {
                        excel[rowIndex, index].Value = item.Action(cellData);
                    }
            }
            var dictWidth = new Dictionary<int, int>();
            foreach (var (item, index) in FieldOption.Select((item, i) => (item, i)))
            {
                dictWidth.Add(index, item.Width);
            }
            return excel.GetStream(dictWidth,autofit);
        }

        /// <summary>
        ///     数据实体导出为Excel
        /// </summary>
        public Stream Export(IEnumerable<T>? data = null,bool autofit = false)
        {
            using var stream = new MemoryStream();
            return Export(stream, data,autofit);
        }

        /// <summary>
        ///     数据实体导出为Excel
        /// </summary>
        public Stream Export(string path, IEnumerable<T>? data = null,bool autofit = false)
        {
            using var fs = new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite);
            return Export(fs, data,autofit);
        }

        /// <summary>
        ///     数据实体导出为Excel
        /// </summary>
        public Stream Export(Stream stream, IEnumerable<T>? data = null,bool autofit = false)
        {
            using var read = ExcelType.CellReaderFunc(stream);
            var exportStream = Export(read, data);
            var dictWidth = new Dictionary<int, int>();
            foreach (var (item, index) in FieldOption.Select((item, i) => (item, i)))
            {
                dictWidth.Add(index, item.Width);
            }
            read.Save(dictWidth,autofit);
            return exportStream;
        }

        #endregion
    }
}