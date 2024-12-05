namespace ExcelSink.Model
{
    internal record MeasureRow(string SensorId, DateOnly Date, MyTimeOnly Time, decimal Power, decimal Temperature);
}
