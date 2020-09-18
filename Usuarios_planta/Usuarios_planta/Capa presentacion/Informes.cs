using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using System.IO;
using Usuarios_planta.Capa_presentacion;

namespace Usuarios_planta.Formularios
{
    public partial class Informes : Form
    {

        MySqlConnection con = new MySqlConnection("server=82.2.121.99;Uid=userapp;password=userapp;database=dblibranza;port=3306;persistsecurityinfo=True;");

        Comandos cmds = new Comandos();
        dia_dia cmds_dia = new dia_dia();

        public Informes()
        {
            InitializeComponent();
        }

        private void Informes_Load(object sender, EventArgs e)
        {

        }

        private void Btn_busqueda_Click(object sender, EventArgs e)
        {
            cmds_dia.busqueda_plano(dgv_datos_plano, Txtbusqueda);
        }
    }
}
