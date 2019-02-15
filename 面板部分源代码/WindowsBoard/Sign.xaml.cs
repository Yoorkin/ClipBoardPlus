using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace WindowsBoard
{
    /// <summary>
    /// Sign.xaml 的交互逻辑
    /// </summary>
    public partial class Sign : Window
    {
        public Sign()
        {
            InitializeComponent();
        }
        public static bool Enable = true;

        public static void ShowInfo(System.Drawing.Color BackGround, BitmapImage Icon)
        {
            if (!Enable) return;
            Sign sign = new Sign();

            sign.Left = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width - 200;
            sign.Top = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height-180;
            SignInfo info = new SignInfo { color = new SolidColorBrush(System.Windows.Media.Color.FromArgb(BackGround.A,BackGround.R,BackGround.G,BackGround.B)) , icon = Icon };
            sign.DataContext = info;
            sign.PlayAni();
        }
        public void PlayAni()
        {
            Topmost = true;


            DoubleAnimation Y = new DoubleAnimation();
            DoubleAnimation Alpha = new DoubleAnimation();
            Alpha.From = 0;
            Alpha.To = 1;
            Y.From = 0;
            Y.By = 50;

            Y.Duration = Alpha.Duration = new Duration(TimeSpan.FromMilliseconds(200)); ;
            this.RenderTransform = new TranslateTransform();
            this.RenderTransform.BeginAnimation(TranslateTransform.YProperty, Y);
            this.grid.BeginAnimation(Grid.OpacityProperty, Alpha);
            this.Show();

            System.Windows.Forms.Timer StartWait = new System.Windows.Forms.Timer();
            StartWait.Interval = 400;
            StartWait.Tick += StartWait_Tick;
            StartWait.Start();
        }

        private void StartWait_Tick(object sender, EventArgs e)
        {
            System.Windows.Forms.Timer timer = sender as System.Windows.Forms.Timer;
            timer.Enabled = false;
            System.Windows.Forms.Timer CloseWait = new System.Windows.Forms.Timer();
            CloseWait.Interval = 200;
            CloseWait.Tick += CloseWait_Tick; ;
            CloseWait.Start();
            DoubleAnimation Y = new DoubleAnimation();
            DoubleAnimation Alpha = new DoubleAnimation();
            Alpha.From = 1;
            Alpha.To = 0;
            Y.From = 50;
            Y.By = 70;

            Y.Duration = Alpha.Duration =  new Duration(TimeSpan.FromMilliseconds(200));;
            this.RenderTransform = new TranslateTransform();
            this.RenderTransform.BeginAnimation(TranslateTransform.YProperty, Y);
            this.grid.BeginAnimation(Grid.OpacityProperty, Alpha);
        }

        private void CloseWait_Tick(object sender, EventArgs e)
        {
            System.Windows.Forms.Timer timer = sender as System.Windows.Forms.Timer;
            timer.Enabled = false;
            this.Close();
        }
    }
    public class SignInfo
    {
        public Brush color { get; set; }
        public BitmapImage icon { get; set; }
    }
}
