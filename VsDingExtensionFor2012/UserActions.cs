using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VitaliiGanzha.VsDingExtension
{
    public partial class UserActions : ToolWindowPane
    {
        
        public UserActions() : base(null)
        {
            this.Caption = "user Emotions";
            this.BitmapResourceID = 301;
            this.BitmapIndex = 1;
            //InitializeComponent();
            this.Content = new Panel();
        }
    }
}
