using System.ServiceProcess;

namespace AssinaturaService
{
    static class Program
    {

        static void Main()
        {
            // assinatura assinatura = new assinatura();
            // assinatura.AssinarCertificado();

            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[]
             {
                new assinatura()
             };
            ServiceBase.Run(ServicesToRun);
        }
    }
}
