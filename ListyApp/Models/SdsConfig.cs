using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ListyApp.Models
{
    public class SdsConfig
    {
        public SdsConfig()
        {
            ItemName = string.Empty;
            ItemValue = string.Empty;
        }

        public SdsConfig(int id, string itemName, string itemValue, bool editable, int viewable)
        {
            Id = id;
            ItemName = itemName;
            ItemValue = itemValue;
            Editable = editable;
            Viewable = viewable;
        }

        public int Id { get; set; }

        [StringLength(50)]
        [Required]
        public string ItemName { get; set; }
        public string ItemValue { get; set; }
        public bool Editable { get; set; }
        public int Viewable { get; set; }
    }
}
