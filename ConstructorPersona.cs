using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace add
{
    // CLASE PARA CREAR OBJETOS DE TIPO PERSONA, CON UN CONSTRUCTOR QUE RECIBE UN PARAMETRO DE TIPO STRING, Y ESTO ES CAPTURADO POR EL ATRIBUTO "NOMBRE"
    class ConstructorPersona
    {
        public string Nombre;
        public int Edad;

        public ConstructorPersona(string nombre, int edad)
        {
            Nombre = nombre;
            Edad = edad;
        }

    }
    
    //INCIO DE LA APLICACION
    class Program
    {
        static void Main()
        {
            ConstructorPersona p = new ConstructorPersona("Juan", 28); // SE INTANCIA EL CONSTRUCTOR (SE CREA EL OBJETO),
                                                                       // los inputs son capturados por el constructor y asignados como valor a los atributos de la clase

            MessageBox.Show(p.Nombre + p.Edad); // ACCESO A LOS ATRIBUTOS DE LA CLASE (MESSAGEBOX Y CONSOLE.WRITELINE SOLO FUNCIONAN CON CONSOLE PROYECTO, NO CON .NET FRAMEWORK)
        }
    }
}
