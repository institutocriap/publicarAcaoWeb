using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;

namespace criapLibrary.types
{
    public class objColaboradores
    {
        //public objColaboradores();

        public string action { get; set; }
        public string colaboradorID { get; set; }
        public string contaActiva { get; set; }
        public Color cor { get; set; }
        public string email { get; set; }
        public Image foto { get; set; }
        public string nome { get; set; }
        public string rowid { get; set; }
        public string sigla { get; set; }
        public bool tecMoodle { get; set; }
        public string telefone { get; set; }
        public string username { get; set; }
    }
}
