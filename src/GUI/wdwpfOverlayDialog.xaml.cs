namespace ExcelDna_MVVM.GUI
{
    using ExcelDna_MVVM.Utils;
    #region Usings
    using NLog;
    using System;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Input;
    using System.Windows.Interop;
    #endregion

    public interface IOverlayHandleClickAndClose
    {
        Action CloseRequested { set; }
        void OutSideClick();
    }

    /// <summary>
    /// Interaction logic for wdwpfOverlayDialog.xaml
    /// </summary>
    public partial class wdwpfOverlayDialog : Window
    {
        #region Logger
        private static Logger logger = LogManager.GetCurrentClassLogger();
        #endregion

        #region Variables & Properties
        public RelayCommand OutsideAreaClick { get; private set; }

        private UserControl childUserControl = null;
        public UserControl ChildUserControl
        {
            get
            {
                return childUserControl;
            }
            set
            {
                if (childUserControl != value)
                {
                    if (!Canvas.Children.Contains(ChildUserControl))
                        Canvas.Children.Remove(ChildUserControl);

                    childUserControl = value;

                    Canvas.Children.Add(ChildUserControl);

                    ResetChildArea();
                }
            }
        }

        private Action closeRequested;
        public Action CloseRequested { get => closeRequested; set => closeRequested = value; }
        private Point topLeftCorner;
        public Point TopLeftCorner
        {
            get
            {
                return topLeftCorner;
            }
            set
            {
                if (topLeftCorner != value)
                {
                    topLeftCorner = value;
                    ResetChildArea();
                }
            }
        }

        public bool WidthMaximized = false;

        public bool HeightMaximized = false;

        private IntPtr parentHwnd;


        public IntPtr ParentHwnd
        {
            get
            {
                return parentHwnd;
            }
            set
            {
                if (value != parentHwnd)
                {
                    parentHwnd = value;
                    ResetChildArea();
                }
            }
        }
        #endregion

        #region Functions
        private void ResetChildArea()
        {
            if (ChildUserControl == null || parentHwnd == null)
                return;

            try
            {
                double dpiX = 1 / Win32Helper.GetDpiXScale;
                double dpiY = 1 / Win32Helper.GetDpiYScale;
                var rct = Win32Helper.GetParentWindowSize(this, parentHwnd);

                new WindowInteropHelper(this).Owner = parentHwnd;

                double yDiff = 0;

                if (Win32Helper.IsMaximized(parentHwnd))
                {
                    var screen = System.Windows.Forms.Screen.FromHandle(parentHwnd);
                    yDiff = rct.Top - screen.Bounds.Top;
                    this.SourceInitialized += (s, a) => this.WindowState = WindowState.Maximized;
                }

                if (Double.IsInfinity(rct.Left) || Double.IsInfinity(rct.Top) || Double.IsInfinity(rct.Height) || Double.IsInfinity(rct.Width))
                    return;

                this.ChildUserControl.SetValue(Canvas.LeftProperty, (TopLeftCorner.X - rct.Left + yDiff) * dpiX);
                this.ChildUserControl.SetValue(Canvas.TopProperty, (TopLeftCorner.Y - rct.Top + yDiff) * dpiY);

                if (HeightMaximized) this.ChildUserControl.SetValue(Canvas.HeightProperty, (rct.Height + rct.Top - TopLeftCorner.Y + yDiff) * dpiY);
                if (WidthMaximized) this.ChildUserControl.SetValue(Canvas.WidthProperty, (rct.Width + rct.Left - TopLeftCorner.X + yDiff) * dpiX);

                if (ChildUserControl is IOverlayHandleClickAndClose cr)
                {
                    cr.CloseRequested = () => { Close(); };
                }


                rct.Scale(dpiX, dpiY);

                this.Top = rct.Top;
                this.Left = rct.Left;
                this.Height = rct.Height;
                this.Width = rct.Width;
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
        }

        private void OnOutsideAreaClick()
        {
            if (ChildUserControl is IOverlayHandleClickAndClose uc)
            {
                uc.OutSideClick();
            }
            else
            {
                CloseRequested?.Invoke();
            }
        }

        private void HandleEsc(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
                OnOutsideAreaClick();
        }
        #endregion

        #region Constructor
        public wdwpfOverlayDialog()
        {
            closeRequested = () => { Close(); };
            InitializeComponent();

#if !DEBUG
            GrayRectangle.Opacity = 0.08;
#else
            GrayRectangle.Opacity = 0.3;
#endif
            OutsideAreaClick = new RelayCommand((o) =>
            {
                OnOutsideAreaClick();
            });
            this.DataContext = this;

            this.KeyDown += new KeyEventHandler(HandleEsc);
        }
        #endregion
    }
}
