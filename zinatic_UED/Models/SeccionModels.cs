using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace zinatic_UED.Models
{
    public class SeccionModels
    {

        //private int idSeccion;
        private int idSeccion;
        private string titulo;
        private string valorDato;        
        private int numeroOrden;
        private string descripcion;

        public int IdSeccion
        {
            get { return idSeccion; }
            set { idSeccion = value; }
        }
        
        public int NumeroOrden
        {
            get { return numeroOrden; }
            set { numeroOrden = value; }
        }
        public string ValorDato
        {
            get { return valorDato; }
            set { valorDato = value; }
        }
        public string Titulo
        {
            get { return titulo; }
            set { titulo = value; }
        }
        public string Descripcion
        {
            get { return descripcion; }
            set { descripcion = value; }
        }


    }
}