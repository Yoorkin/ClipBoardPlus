using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Runtime.InteropServices;
using System.Windows.Interop;
using System.Diagnostics;
using System.Collections.ObjectModel;
using System.IO;
using System.Threading;
using System.Drawing;
using System.Collections;
using System.Text.RegularExpressions;
using System.Windows.Media.Animation;
using Microsoft.Win32;
using Microsoft;
//using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace WindowsBoard
{

    public class ClipBoardViewer: System.Windows.Forms.Form
    {
        [DllImport("User32.dll")]
        public static extern int SetClipboardViewer(int hWndNewViewer);
        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern bool ChangeClipboardChain(IntPtr hWndRemove,IntPtr hWndNewNext);
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int SendMessage(IntPtr hwnd, int wMsg,IntPtr wParam,IntPtr lParam);

        public delegate void OnClipBoardChangeHander();
        public event OnClipBoardChangeHander OnClipBoardChange;
        public bool Enable;

        IntPtr nextClipboardViewer;
        public MainWindow Board;
        public System.Windows.Forms.NotifyIcon notifyIcon=new System.Windows.Forms.NotifyIcon();
        public ClipBoardViewer(Window window)
        {
            
            Board = window as MainWindow;

            //设置剪贴板监视
            nextClipboardViewer = (IntPtr)SetClipboardViewer((int)this.Handle);

            //设置托盘图标
            notifyIcon.Visible = true;
            notifyIcon.Text = "Windows Board正收集剪贴板数据";
            notifyIcon.Icon =  new System.Drawing.Icon(Application.GetResourceStream(new Uri("pack://application:,,,/Image/clipboardicon.ico")).Stream,new System.Drawing.Size(106,128));//Environment.CurrentDirectory + "\\Img\\icon.ico"); //System.Drawing.Icon.ExtractAssociatedIcon(System.Windows.Forms.Application.ExecutablePath);
            notifyIcon.MouseClick += NotifyIcon_MouseClick;

            //设置托盘菜单
            System.Windows.Forms.MenuItem Exit = new System.Windows.Forms.MenuItem("退出",Exit_OnClick);
            System.Windows.Forms.MenuItem AutoRun = new System.Windows.Forms.MenuItem("开机时启动",AutoRun_Onclick);
            System.Windows.Forms.MenuItem ShowSign = new System.Windows.Forms.MenuItem("显示气泡", ShowSign_Onclick);
            System.Windows.Forms.MenuItem ViewEnable = new System.Windows.Forms.MenuItem("开启剪贴板监视", ViewEnable_Onclick);
            System.Windows.Forms.MenuItem AddHelp = new System.Windows.Forms.MenuItem("显示教程", AddHelp_Onclick);
            Enable = ViewEnable.Checked = true;

            //如果未开启监视则显示暂停监视横幅
            Board.ViewerStateTip.Visibility = ViewEnable.Checked ? Visibility.Hidden : Visibility.Visible;

            //加载注册表
            AppRegistry AppReg = new AppRegistry("ClipBoardPlus", AppDomain.CurrentDomain.BaseDirectory + "/ClipBoardPlus.exe");
            ShowSign.Checked = (bool)AppReg.Get("ShowSign");
            if ((string)AppReg.Get("FirstRun","True")=="True") AppReg.AutoRun = true;
            AutoRun.Checked = AppReg.AutoRun;
            notifyIcon.ContextMenu = new System.Windows.Forms.ContextMenu(new System.Windows.Forms.MenuItem[] { AddHelp,AutoRun, ShowSign, ViewEnable ,Exit });

        }
        private void AddHelp_Onclick(object sender, EventArgs e)
        {
            Board.AddHelpToTimeLine();
            Board.Show();
            Board.Activate();
        }

        private void ViewEnable_Onclick(object sender, EventArgs e)
        {
            System.Windows.Forms.MenuItem item = sender as System.Windows.Forms.MenuItem;
            Enable = item.Checked = !item.Checked;
            Board.ViewerStateTip.Visibility = item.Checked ? Visibility.Hidden : Visibility.Visible;
        }

        private void ShowSign_Onclick(object sender, EventArgs e)
        {
            System.Windows.Forms.MenuItem item = sender as System.Windows.Forms.MenuItem;
            Sign.Enable = item.Checked = !item.Checked;
            AppRegedit.Set("ShowSign", Sign.Enable);
        }

        private void AutoRun_Onclick(object sender, EventArgs e)
        {
            System.Windows.Forms.MenuItem item = sender as System.Windows.Forms.MenuItem;
            AppRegedit.AutoRun = item.Checked = !item.Checked;
        }

        private void Exit_OnClick(object sender, EventArgs e)
        {
            App.Current.Shutdown();
        }



        private void NotifyIcon_MouseClick(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            if(e.Button==System.Windows.Forms.MouseButtons.Left)
            {
               Board.Show();
               Board.Activate();
            }
        }

        protected override void WndProc(ref System.Windows.Forms.Message m)
        {
            // defined in winuser.h
            const int WM_DRAWCLIPBOARD = 0x308;
            const int WM_CHANGECBCHAIN = 0x030D;

            switch (m.Msg)
            {
                case WM_DRAWCLIPBOARD:
                    if(Enable&&OnClipBoardChange!=null) OnClipBoardChange();
                    SendMessage(nextClipboardViewer, m.Msg, m.WParam, m.LParam);
                    break;

                case WM_CHANGECBCHAIN:
                    if (m.WParam == nextClipboardViewer)
                        nextClipboardViewer = m.LParam;
                    else
                        SendMessage(nextClipboardViewer, m.Msg, m.WParam,
                                    m.LParam);
                    break;

                default:
                    base.WndProc(ref m);
                    break;
            }
        }
        protected override void Dispose(bool disposing)
        {
            ChangeClipboardChain(this.Handle, nextClipboardViewer);
            base.Dispose(disposing);
        }
        //Code: https://stackoverflow.com/questions/24863594/faster-way-to-convert-bitmapsource-to-bitmapimage#

    }
    public class TopMostTool
    {
        //代码来自 http://blog.csdn.net/u014434080/article/details/51029959

        public static int SW_SHOW = 5;
        public static int SW_NORMAL = 1;
        public static int SW_MAX = 3;
        public static int SW_HIDE = 0;
        public static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);    //窗体置顶
        public static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);    //取消窗体置顶
        public const uint SWP_NOMOVE = 0x0002;    //不调整窗体位置
        public const uint SWP_NOSIZE = 0x0001;    //不调整窗体大小
        public bool isFirst = true;

        [DllImport("user32.dll")]
        public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint uFlags);

        [DllImport("user32.dll", EntryPoint = "ShowWindow")]
        public static extern bool ShowWindow(System.IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        /// <summary>
        /// 在外面的方法中掉用这个方法就可以让浮动条（CustomBar）始终置顶
        /// CustomBar是我的程序中需要置顶的窗体的名字，你们可以根据需要传入不同的值
        /// </summary>
        public static void SetTopCustomBar(string WindowName)
        {
            IntPtr CustomBar = FindWindow(null, WindowName);    //CustomBar是我的程序中需要置顶的窗体的名字
            if (CustomBar != null)
            {
                SetWindowPos(CustomBar,HWND_TOPMOST, 0, 0, 0, 0,SWP_NOMOVE | SWP_NOSIZE);
            }
        }

    }
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    ///
    public partial class MainWindow : Window
    {
        public ObservableCollection<ClipBoardItem> Item = new ObservableCollection<ClipBoardItem>();
        public bool IgnoreNextChange = false;
        public bool ScreenShotEnd = true, HasScreenPicture = false;
        public string DataFile = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\ClipBoardPlus";
        public MainWindow()
        {
            InitializeComponent();
            this.Hide();
            this.Top = 0;
            this.Left = System.Windows.Forms.SystemInformation.WorkingArea.Width - this.Width;
            this.Height = System.Windows.Forms.SystemInformation.WorkingArea.Height;
            Timeline.Width = this.Width+10;
            Timeline.Height = this.Height-50;
            ClipBoardViewer viewer = new ClipBoardViewer(this);
            viewer.OnClipBoardChange += Viewer_OnClipBoardChange;
            Timeline.Items.Clear();
            Timeline.ItemsSource = Item;
            Timeline.SelectedValuePath = "Data";
        }
         
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            TopMostTool.SetTopCustomBar(this.Title);
            Clean.ImageSource = new BitmapImage(new Uri("pack://application:,,,/Image/Clean.png"));
            ScreenShoot.ImageSource = new BitmapImage(new Uri("pack://application:,,,/Image/ScreenShot.png"));
        }

        public void Viewer_OnClipBoardChange()
        {
            bool HasPopSign = false;
            ClipBoardItem item = new ClipBoardItem();
            item.Title = "剪贴板 " + DateTime.Now.Hour.ToString() + ":" + DateTime.Now.Minute.ToString();
            System.Windows.IDataObject data =System.Windows.Clipboard.GetDataObject();

            string[] Formats = data.GetFormats();//将剪贴板所有东西取出封入Item.Data
            if (Formats.Length == 0) return;
            if (IgnoreNextChange) { IgnoreNextChange = false;return; }

            try
            {
                foreach (string Format in Formats)
                {
                    object obj = data.GetData(Format);
                    if (obj != null) item.Data.Add(Format, obj);
                }
            }
            catch
            {
                //无法识别的格式
            }

            try
            {
                if (data.GetDataPresent("FileDrop"))
                {

                    string[] Drops = data.GetData("FileDrop") as string[];
                    if (!Clipboard.ContainsImage())
                    {
                        item.Title = "文件 " + DateTime.Now.Hour + ":" + DateTime.Now.Minute;
                        ObservableCollection<IconFile> list = new ObservableCollection<IconFile>();
                        foreach (string Drop in Drops)
                        {
                            IconFile file = new IconFile();
                            file.FileName = Drop;
                            try
                            {
                                System.Drawing.Image img = (System.Drawing.Image)System.Drawing.Icon.ExtractAssociatedIcon(Drop).ToBitmap();
                                if (img != null) file.Icon = new Bitmap(img).ToBitmapImage(); else file.Icon = ClipBoardItem.IconFile;
                            }
                            catch
                            {

                            }
                            list.Add(file);
                        }
                        item.IconFilelist = list;
                        item.SetFrame(ClipBoardItem.FrameStyle.FileAndIcon);
                        item.Icon = ClipBoardItem.IconFile;
                    }
                    else
                    {
                        item.Img = new BitmapImage(new Uri(Drops[0]));
                        item.SetFrame(ClipBoardItem.FrameStyle.FullImage);
                        item.Icon = ClipBoardItem.IconClipBoard;
                    }

                }
                if (data.GetDataPresent(typeof(BitmapSource)))
                {
                    BitmapSource bitmap = data.GetData(typeof(BitmapSource)) as BitmapSource;
                    item.Img = bitmap.ToBitmapImage();
                    item.SetFrame(ClipBoardItem.FrameStyle.FullImage);
                    item.Icon = ClipBoardItem.IconClipBoard;
                }
                if (data.GetDataPresent(typeof(Bitmap)))
                {
                    Bitmap bitmap = data.GetData(typeof(Bitmap)) as Bitmap;
                    item.Img = bitmap.ToBitmapImage();
                    item.SetFrame(ClipBoardItem.FrameStyle.FullImage);
                    item.Icon = ClipBoardItem.IconClipBoard;
                }
                if (data.GetDataPresent(typeof(string)))
                {
                    item.Text = (string)data.GetData(typeof(string));
                    if (item.Text != "") item.FullImage = Visibility.Hidden;
                    item.SetFrame(ClipBoardItem.FrameStyle.ImageAndText);
                    item.Icon = ClipBoardItem.IconClipBoard;
                    if(!data.GetDataPresent(DataFormats.Html))
                    {
                        if(item.Text.Trim(' ').Substring(0,7).ToUpper()=="HTTP://"|| item.Text.Trim(' ').Substring(0, 8).ToUpper() == "HTTPS://")
                        {
                            item.SetFrame(ClipBoardItem.FrameStyle.Web);
                            item.Icon = ClipBoardItem.IconWeb;
                            item.Title= "网址 " + DateTime.Now.Hour.ToString() + ":" + DateTime.Now.Minute.ToString();
                        }
                    }
                }
                if (data.GetDataPresent(DataFormats.Html))
                {
                    string Html = data.GetData(DataFormats.Html) as string;
                    Match match = Regex.Match(Html, ".:([^/*]).*?((jpg)|(png)|(jpeg)|(gif)|(bmp))");
                    item.Img = new BitmapImage(new Uri(match.Value));
                    item.SetFrame(ClipBoardItem.FrameStyle.ImageAndText);
                    item.Icon = ClipBoardItem.IconClipBoard;
                }

                if (item.Web == Visibility.Hidden && item.ImageAndText == Visibility.Hidden && item.FileAndIcon==Visibility.Hidden && item.FullImage==Visibility.Hidden )
                {
                    Bitmap screen = new Bitmap(900,600);
                    Graphics Graphics = Graphics.FromImage(screen);
                    Graphics.CopyFromScreen(new System.Drawing.Point(System.Windows.Forms.Control.MousePosition.X - 450, System.Windows.Forms.Control.MousePosition.Y - 300), new System.Drawing.Point(0, 0), new System.Drawing.Size(900, 600));
                    item.Img = screen.ToBitmapImage();
                    item.Title = "未知数据 "+DateTime.Now.Hour.ToString() + ":" + DateTime.Now.Minute.ToString();
                    item.Text = "无法预览数据,已显示复制时鼠标周围的场景";
                    item.Icon = ClipBoardItem.IconUnknown;
                    item.ImageAndText = Visibility.Visible;
                    Sign.ShowInfo(System.Drawing.Color.Orange, ClipBoardItem.IconUnknown);
                    HasPopSign = true;
                }
            }
            catch
            {

            }

            if(!HasPopSign)
            {
                if(item.ImageAndText==Visibility.Visible||item.FullImage==Visibility.Visible) Sign.ShowInfo(System.Drawing.Color.DeepSkyBlue, ClipBoardItem.IconClipBoard);
                if(item.FileAndIcon==Visibility.Visible) Sign.ShowInfo(System.Drawing.Color.YellowGreen, ClipBoardItem.IconFile);
                if(item.Web==Visibility.Visible) Sign.ShowInfo(System.Drawing.Color.Purple, ClipBoardItem.IconWeb);
            }
            Item.Insert(0,item);

        }
        

        private void Window_Deactivated(object sender, EventArgs e)
        {
            this.RenderTransform = new TranslateTransform();
            Storyboard storyboard = new Storyboard();
            DoubleAnimationUsingKeyFrames Frame = new DoubleAnimationUsingKeyFrames();
            EasingDoubleKeyFrame Key1 = new EasingDoubleKeyFrame(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width, KeyTime.FromTimeSpan(TimeSpan.FromMilliseconds(100)))
            { EasingFunction = new CircleEase() { EasingMode = EasingMode.EaseOut } };
            EasingDoubleKeyFrame Key2 = new EasingDoubleKeyFrame(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width - 400, KeyTime.FromTimeSpan(TimeSpan.FromMilliseconds(0)))
            { EasingFunction = new CircleEase() { EasingMode = EasingMode.EaseOut } };
            Frame.KeyFrames.Add(Key2);
            Frame.KeyFrames.Add(Key1);
            storyboard.Children.Add(Frame);
            Storyboard.SetTarget(Frame, this);
            Storyboard.SetTargetProperty(Frame, new PropertyPath("(Window.Left)"));
            storyboard.Completed += Storyboard_Completed;
            storyboard.Begin();
        }
        private void Storyboard_Completed(object sender, EventArgs e)
        {
            this.Hide();
        }


        private void Pop_Click(object sender, RoutedEventArgs e)
        {
            IgnoreNextChange = true;
            Button btn = sender as Button;
            ClipBoardItem item = (ClipBoardItem)btn.DataContext;
            Clipboard.Clear();
            IDataObject dataObject = new DataObject();
            foreach(string Format in item.Data.Keys)
                dataObject.SetData(Format, item.Data[Format]);

            Clipboard.SetDataObject(dataObject);
            Item.Remove(item);
            this.Hide();
        }

        private void Pow_Click(object sender, RoutedEventArgs e)
        {
            IgnoreNextChange = true;
            Button btn = sender as Button;
            ClipBoardItem item = (ClipBoardItem)btn.DataContext;
            Clipboard.Clear();
            IDataObject dataObject = new DataObject();
            foreach (string Format in item.Data.Keys)
                dataObject.SetData(Format, item.Data[Format]);

            Clipboard.SetDataObject(dataObject);
            this.Hide();
        }

        private void Remove_Click(object sender, RoutedEventArgs e)
        {
 
            Button btn = sender as Button;
            Item.Remove((ClipBoardItem)btn.DataContext);
        }



        private void Window_Activated(object sender, EventArgs e)
        {
            if (Item.Count > 0) Clean.IsEnabled = true;
            this.RenderTransform = new TranslateTransform();
            Storyboard storyboard = new Storyboard();
            DoubleAnimationUsingKeyFrames Frame = new DoubleAnimationUsingKeyFrames();
            EasingDoubleKeyFrame Key1 = new EasingDoubleKeyFrame();
            Key1.Value = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width;
            Key1.KeyTime = KeyTime.FromTimeSpan(TimeSpan.FromMilliseconds(0));
            Key1.EasingFunction = new CircleEase() { EasingMode = EasingMode.EaseOut };
            EasingDoubleKeyFrame Key2 = new EasingDoubleKeyFrame();
            Key2.Value = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width - 400;
            Key2.KeyTime = KeyTime.FromTimeSpan(TimeSpan.FromMilliseconds(200));
            Key2.EasingFunction = new CircleEase() { EasingMode = EasingMode.EaseOut };
            Frame.KeyFrames.Add(Key1);
            Frame.KeyFrames.Add(Key2);
            storyboard.Children.Add(Frame);
            Storyboard.SetTarget(Frame, this);
            Storyboard.SetTargetProperty(Frame, new PropertyPath("(Window.Left)"));
            storyboard.Begin();

        }



        private void Clean_Click(object sender, RoutedEventArgs e)
        {
            ConfirmCleanBtn.Visibility = Visibility.Visible;
            ConfirmClean.Focus();
        }
        private void ConfirmClean_Click(object sender, RoutedEventArgs e)
        {
            Item.Clear();
            Clean.IsEnabled = false;
            ConfirmCleanBtn.Visibility = Visibility.Hidden;
        }
        private void ConfirmClean_LostFocus(object sender, RoutedEventArgs e)
        {
            ConfirmCleanBtn.Visibility = Visibility.Hidden;
        }

        private void ImageListBox_Loaded(object sender, RoutedEventArgs e)
        {
            ListBox list = sender as ListBox;
            ClipBoardItem item = (ClipBoardItem)list.DataContext;
            list.ItemsSource = item.IconFilelist;
        }

        private void ScreenShoot_Click(object sender, RoutedEventArgs e)
        {
            ScreenShotEnd = false;
            //Bitmap screen = new Bitmap(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width,System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height);
            //Graphics Graphics = Graphics.FromImage(screen);
            this.Hide();
            //Graphics.CopyFromScreen(new System.Drawing.Point(0,0), new System.Drawing.Point(0, 0), new System.Drawing.Size(screen.Width , screen.Height ));
            Process Shot = new Process();
            Shot.StartInfo.FileName = AppDomain.CurrentDomain.BaseDirectory + "\\ClipBoardData\\ScreenShot.exe";
            Shot.Start();
            Shot.EnableRaisingEvents = true;
            Shot.Exited += Shot_Exited;
            do { } while (ScreenShotEnd != true);
            ClipBoardItem item = new WindowsBoard.ClipBoardItem();
            if (File.Exists(AppDomain.CurrentDomain.BaseDirectory + "\\ClipBoardData\\Tem.png"))
            {
                DirectoryInfo Dir = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory + "\\ClipBoardData\\ScreenShotData");
                FileInfo[] ListFile = Dir.GetFiles();
                int Count = ListFile.Length + 1;
                File.Move(AppDomain.CurrentDomain.BaseDirectory + "\\ClipBoardData\\Tem.png", AppDomain.CurrentDomain.BaseDirectory + "\\ClipBoardData\\ScreenShotData\\Tem" +Count + ".png");
                //File.Delete(AppDomain.CurrentDomain.BaseDirectory + "\\ClipBoardData\\Tem.png");
                item.Img = new BitmapImage(new Uri( AppDomain.CurrentDomain.BaseDirectory + "\\ClipBoardData\\ScreenShotData\\Tem" + Count +".png"));
                item.Title = "屏幕截图 " + DateTime.Now.Hour.ToString() + ":" + DateTime.Now.Minute.ToString();
                item.FullImage = Visibility.Visible;
                item.Icon = ClipBoardItem.IconScreenShot;
                Item.Insert(0,item);
                item.Data.Add(DataFormats.Bitmap, item.Img);
            }

        }

        private void Shot_Exited(object sender, EventArgs e)
        {
            ScreenShotEnd = true;
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {

        }

        private void OpenWeb_Click(object sender, RoutedEventArgs e)
        {
            Button btn = sender as Button;
            ClipBoardItem item = (ClipBoardItem)btn.DataContext;
            System.Diagnostics.Process.Start(item.Text);
        }

        public void AddHelpToTimeLine()
        {
            Item.Add(new ClipBoardItem() { FullImage = Visibility.Visible, Img = new BitmapImage(new Uri("Pack://application:,,,/Image/Welcome/Banner.png")) });
            Item.Add(new ClipBoardItem() { ImageAndText = Visibility.Visible, Title = "如何使用ClipBoard+?", Text = "阅读以下条目了解功能" + Environment.NewLine, Icon = ClipBoardItem.IconUnknown });
            Item.Add(new ClipBoardItem() { ImageAndText = Visibility.Visible, Title = "保存你的剪贴板历史信息", Text = " · 当你复制数据时,ClipBoard+会弹出气泡提示数据已保存\n · 点击右下方托盘区图标唤出此面板", Img = new BitmapImage(new Uri("Pack://application:,,,/Image/Welcome/Sign2.png")), Icon = ClipBoardItem.IconClipBoard });
            Item.Add(new ClipBoardItem() { ImageAndText = Visibility.Visible, Title = "编辑历史数据",Text= " · 鼠标移入时,卡片上方将出现以上三个选项\n",Icon=ClipBoardItem.IconClipBoard, Img = new BitmapImage(new Uri("Pack://application:,,,/Image/Welcome/3Btn.png"))});
            Item.Add(new ClipBoardItem() { ImageAndText = Visibility.Visible, Title = "支持预览复制的文件", Text = " · ClipBoard+会识别复制的文件,并将图标和路径在卡片上列出", Icon = ClipBoardItem.IconFile, Img = new BitmapImage(new Uri("Pack://application:,,,/Image/Welcome/fle.png")) });

            ClipBoardItem file = new ClipBoardItem();
            file.FileAndIcon = Visibility.Visible;
            file.Icon = ClipBoardItem.IconFile;
            file.Title = "示例";
            file.IconFilelist = new ObservableCollection<IconFile>()
            {
                new IconFile() { Icon = ClipBoardItem.IconFile, FileName = "C:/Program Files/ClipBoardPlus/示例文件夹/示例文件1.txt" },
                new IconFile() { Icon = ClipBoardItem.IconFile, FileName = "C:/Program Files/ClipBoardPlus/示例文件夹/示例文件2.txt" },
                new IconFile() { Icon = ClipBoardItem.IconFile, FileName = "C:/Program Files/ClipBoardPlus/示例文件夹/示例文件3.txt" },
                new IconFile() { Icon = ClipBoardItem.IconFile, FileName = "C:/Program Files/ClipBoardPlus/示例文件夹/示例文件4.txt" },
                new IconFile() { Icon = ClipBoardItem.IconFile, FileName = "C:/Program Files/ClipBoardPlus/示例文件夹/示例文件5.txt" },
                new IconFile() { Icon = ClipBoardItem.IconFile, FileName = "C:/Program Files/ClipBoardPlus/示例文件夹/示例文件6.txt" }
            };
            Item.Add(file);

            Item.Add(new ClipBoardItem() { ImageAndText = Visibility.Visible, Title = "识别网址", Text = " · 复制正确的单条网址时,卡片左方将出现打开网页按钮。\n · 试试这个:", Icon = ClipBoardItem.IconWeb, Img = new BitmapImage(new Uri("Pack://application:,,,/Image/Welcome/htp.png")) });
            Item.Add(new ClipBoardItem() { Web = Visibility.Visible, Title = "必应搜索", Text = "http://cn.bing.com/", Icon = ClipBoardItem.IconWeb });
            Item.Add(new ClipBoardItem() { ImageAndText = Visibility.Visible, Title = "快速截图", Text = " · 截图会保存在面板中\n", Icon = ClipBoardItem.IconScreenShot, Img = new BitmapImage(new Uri("Pack://application:,,,/Image/Welcome/sht.png")) });
            Item.Add(new ClipBoardItem() { ImageAndText = Visibility.Visible, Title = "清空列表", Text = " · 已了解全部功能?\n · 双击右上角按钮清空列表,体验ClipBoard+带来的便利。", Icon = ClipBoardItem.IconClipBoard,Img = new BitmapImage(new Uri("Pack://application:,,,/Image/Welcome/cls.png")) });

        }
    }

    public class ClipBoardItem
    {
        public ClipBoardItem ()
        {
            Data = new Hashtable();
            Web = FileAndIcon = ImageAndText = FullImage = Visibility.Hidden;
        }
        public static BitmapImage IconUnknown = new BitmapImage(new Uri("pack://application:,,,/Image/Unknown.png"));
        public static BitmapImage IconFile = new BitmapImage(new Uri("pack://application:,,,/Image/File.png"));
        public static BitmapImage IconClipBoard = new BitmapImage(new Uri("pack://application:,,,/Image/ClipBoard.png"));
        public static BitmapImage IconScreenShot = new BitmapImage(new Uri("pack://application:,,,/Image/ScreenShot.png"));
        public static BitmapImage IconWeb = new BitmapImage(new Uri("pack://application:,,,/Image/Web.png"));

        public BitmapImage Icon { get; set; }
        public ObservableCollection<IconFile> IconFilelist { get; set; }
        public string Title { get; set; }
        public string Text { get; set; }
        public BitmapImage Img { get; set; }
        public Hashtable Data { get; set; }
        public Visibility FullImage { get; set; }
        public Visibility ImageAndText { get; set; }
        public Visibility FileAndIcon { get; set; }
        public Visibility Web { get; set; }

        public enum FrameStyle {FullImage,ImageAndText,FileAndIcon,Web}
        public void SetFrame(FrameStyle style)
        {
            FileAndIcon = ImageAndText = FullImage = Visibility.Hidden;
            switch(style)
            {
                case FrameStyle.FullImage:FullImage = Visibility.Visible; break;
                case FrameStyle.FileAndIcon:FileAndIcon = Visibility.Visible;break;
                case FrameStyle.ImageAndText:ImageAndText = Visibility.Visible;break;
                case FrameStyle.Web:Web = Visibility.Visible;break;
            }
        }
    }
    public class IconFile
    {
        public string FileName { get; set; }
        public BitmapImage Icon { get; set; }
    }

    public class MyImageButton :Button
    {
        public readonly static DependencyProperty ImageSourceProperty = DependencyProperty.Register("ImageSource", typeof(BitmapImage), typeof(MyImageButton));
        public BitmapImage ImageSource { get { return (BitmapImage)GetValue(ImageSourceProperty); } set { SetValue(ImageSourceProperty, value); } }
    }

    public static class Tool
    {
        internal static BitmapImage ToBitmapImage(this BitmapSource bitmapSource)
        {
            JpegBitmapEncoder encoder = new JpegBitmapEncoder();
            MemoryStream memorystream = new MemoryStream();
            BitmapImage tmpImage = new BitmapImage();
            encoder.Frames.Add(BitmapFrame.Create(bitmapSource));
            encoder.Save(memorystream);

            tmpImage.BeginInit();
            tmpImage.StreamSource = new MemoryStream(memorystream.ToArray());
            tmpImage.EndInit();

            memorystream.Close();
            return tmpImage;
        }
        internal static BitmapImage ToBitmapImage(this Bitmap bitmap)
        {
            Bitmap bitmapSource = new Bitmap(bitmap.Width, bitmap.Height);
            int i, j;
            for (i = 0; i < bitmap.Width; i++)
                for (j = 0; j < bitmap.Height; j++)
                {
                    System.Drawing.Color pixelColor = bitmap.GetPixel(i, j);
                    System.Drawing.Color newColor = System.Drawing.Color.FromArgb(pixelColor.A,pixelColor.R, pixelColor.G, pixelColor.B);
                    bitmapSource.SetPixel(i, j, newColor);
                }
            MemoryStream ms = new MemoryStream();
            bitmapSource.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
            BitmapImage bitmapImage = new BitmapImage();
            bitmapImage.BeginInit();
            bitmapImage.StreamSource = new MemoryStream(ms.ToArray());
            bitmapImage.EndInit();

            return bitmapImage;
        }
        //========================================================================================================


    }
 
}


