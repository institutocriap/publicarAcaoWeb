using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace criapLibrary.types
{
    public class objDestacamento
    {
        //public objDestacamento();

        public string codigo_Colaborador { get; set; }
        public string codigo_Google { get; set; }
        public string codigo_Validacao { get; set; }
        public int colaboradorID { get; set; }
        public DateTime data_Fim { get; set; }
        public DateTime data_Inicio { get; set; }
        public string descricao { get; set; }
        public int label { get; set; }
        public string local { get; set; }
        public string notas { get; set; }
        public string rowid { get; set; }
        public string rowid_Sessao { get; set; }
        public string rowid_TipoDest { get; set; }
    }
}