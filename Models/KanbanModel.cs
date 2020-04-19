
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace Karban.Models
{
    public class KanbanModel
    {
        public string CodeId { get; set; }
        public string Subject { get; set; }

        [Display(Name ="Developer")]
        public string DeveloperName { get; set; }
        public string AssignedOn { get; set; }
        public string Priority { get; set; }
        public string StatusCode { get; set; }
    }
}
