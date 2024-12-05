namespace ExcelSink.Model
{
    internal record MyTimeOnly(int hour, int minute, int second)
    {
        public override string ToString()
        {
            return $"{hour:00}:{minute:00}:{second:00}";
        }
    }
}
