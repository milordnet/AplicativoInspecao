using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MilordAssinatura
{
    public interface Iassinar
    {

        void Assinar(listaservico listaservico, bool erro);


    }
}
