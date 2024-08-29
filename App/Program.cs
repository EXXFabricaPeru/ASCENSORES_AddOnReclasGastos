using AddOnRclsGastos.App;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AddOnRclsGastos
{
    internal static class Program
    {
        /// <summary>
        /// Punto de entrada principal para la aplicación.
        /// </summary>
        [STAThread]
        static void Main()
        {
            try
            {
                Main oApp = new Main();
                if (Globals.continuar == 0)
                    System.Windows.Forms.Application.Run();
            }
            catch (Exception e)
            {
                Globals.SBO_Application.MessageBox(e.Message.ToString());
            }
        }
    }
}
