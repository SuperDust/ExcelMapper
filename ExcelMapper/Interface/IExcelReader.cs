namespace ExcelMapper.Interface
{
    /// <summary>
    ///     尝试使用 IExcelRead 统一组件的调用
    /// </summary>
    public interface IExcelReader : IExcelReader<string>
    {
    }

    /// <summary>
    ///     尝试使用 IExcelRead 统一组件的调用
    /// </summary>
    public interface IExcelReader<T> : IExcelContainer<T>
    {
    }
}