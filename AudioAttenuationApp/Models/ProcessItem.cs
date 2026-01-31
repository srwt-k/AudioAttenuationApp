using System.Drawing;

namespace AudioAttenuationApp.Models
{
    public class ProcessItem(int id, string name, Icon icon)
    {
        public int Id { get; set; } = id;
        public string Name { get; set; } = name;
        public Icon Icon { get; set;  } = icon;
    }
}
