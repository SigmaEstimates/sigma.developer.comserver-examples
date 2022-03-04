using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Sigma_Integration
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Sigma.ISigmaApplication sigmaApp = new Sigma.SigmaApplication(); //Connect to the Sigma COM Server
            Sigma.ISigmaProject sigmaProject = sigmaApp.Projects.Add(); //Add (create) a new Sigma Project
            sigmaProject.Name = "Sigma Integration Example"; //Set the name of the project
            Sigma.ISigmaProjectPropertySet navSet = sigmaProject.ProjectProperties.Add("Navision"); //Create a new property set named Navision
            Sigma.ISigmaProjectProperty navProjectID = navSet.Add("ProjectID"); //Create a new property in the Navision set named ProjectID
            navProjectID.Value = "123"; //and assign a value to it
            sigmaApp.Activate(); //Show the application
        }
    }
}
