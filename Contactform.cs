using System.Windows.Forms;
using System.Reflection;

namespace PowerpointGenerater2
{
    public partial class Contactform : Form
    {
        public Contactform()
        {
            InitializeComponent();
        }

        private void Contactform_Load(object sender, System.EventArgs e)
        {
            lblBuild.Text = AssemblyName.GetAssemblyName("PowerpointGenerater.exe").Version.ToString();
        }
    }
}
