using System.Drawing;

namespace AudioAttenuationApp.Models
{
    public class ProcessItem(uint id, string name, Icon icon)
    {
        public uint Id { get; set; } = id;
        public string Name { get; set; } = name;
        public Icon Icon { get; set;  } = icon;
    }
}
