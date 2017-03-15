namespace MySQL.Models
{
    public class Dict
    {
        public long? 字段序号 { get; set; }
        public string 字段名 { get; set; }
        public string 标识 { get; set; }
        public string 主键 { get; set; }
        public string 数据类型 { get; set; }
        public long? 占用字节数 { get; set; }
        public long? 长度 { get; set; }
        public long? 小数位数 { get; set; }
        public string 允许空 { get; set; }
        public string 默认值 { get; set; }
        public string 字段说明 { get; set; }
    }
}
