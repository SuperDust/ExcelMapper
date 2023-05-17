// See https://aka.ms/new-console-template for more information

using ExcelMapper;
using ExcelMapper.Attributes;

// 导出
var stream = await ExportConfig<object>.GenDefaultConfig(new List<object>
    {
        new
        {
            Name = "张三",
            Sex = 0,
            Birthday =  DateTime.Now
        },
        new
        {
            Name = "李四",
            Sex = 1,
            Birthday =  DateTime.Now
        }
    })
    .Add("名称", "Name", 30)
    .Add("性别", "Sex", 40)
    .Add("生日", "Birthday", 50)
    .Handler("性别", t => ((dynamic)t).Sex == 1 ? "女" : "男")
    .Handler("生日", t => ((dynamic)t).Birthday.ToString("yyyy-MM-dd HH:mm:ss"))
    .ExportAsync();
await using var sw = File.Create("D:\\monitor\\test.xlsx");
await stream.CopyToAsync(sw);

// 导入
var result = await ReadConfig<Demo>.ExcelToEntityAsync(stream);


public class Demo
{

    [Excel(ExcelField = "名称")]
    public string Name { get; set; }
    [Excel(ExcelField = "性别", ReadConverterExp = "0=男,1=女", Separator = ",")]
    public int Sex { get; set; }
    [Excel(ExcelField = "生日")]
    public DateTime? Birthday { get; set; }
}