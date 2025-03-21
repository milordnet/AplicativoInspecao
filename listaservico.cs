using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MilordAssinatura
{
    public class listaservico
    {

        public listaservico()
        {
            List<servicoagrupado> srvagrupado = new List<servicoagrupado>();
            agrupado = srvagrupado;
        }

        public ServicosAssinatura ServicoSelecionado { get; set; }

        public EventLog eventLog1 { get; set; }


        public List<servicoagrupado> agrupado { get; set; }

        public class servicoagrupado
        {

            public servicoagrupado()
            {
                List<item> _item = new List<item>();
                items = _item;
            }

            public List<item> items { get; set; }
            public int CodigoGerente { get; set; }
            public string HoraAssinatura { get; set; }
            public DateTime DataAssinatura { get; set; }


            public class item
            {

                public int CodigoServico { get; set; }

            }
        }

    }
}
