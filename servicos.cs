using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MilordAssinatura
{
    public class servicos : List<servico>
    {
        public servicos()
        {
        }

        public Iassinar ServicoAssinar { get; set; }


    }
}
