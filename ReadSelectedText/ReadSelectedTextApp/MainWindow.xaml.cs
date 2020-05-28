using ReadSelectedText;
using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;
using System.Speech.Synthesis;

namespace ReadSelectedTextApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private WindowInteropHelper nativeWindow;
        private SelectionReader selectionReader;
        private IntPtr environmentWindow;

        public MainWindow()
        {
            string activeWindow = Environment.GetEnvironmentVariable("GPII_ACTIVE_WINDOW");
            if (!string.IsNullOrEmpty(activeWindow) && int.TryParse(activeWindow, out int n))
            {
                environmentWindow = new IntPtr(n);
            }

            if (environmentWindow != IntPtr.Zero)
            {
                this.Visibility = Visibility.Hidden;
                this.WindowState = WindowState.Minimized;
            }

            InitializeComponent();
        }

        private async void Button_Click(object sender, RoutedEventArgs e)
        {
            await this.SaySelection();
        }

        private async Task SaySelection(IntPtr? activeWindow = null)
        {
            string text = await this.selectionReader.GetSelectedText(activeWindow);
            SpeechSynthesizer speech = new SpeechSynthesizer();
            speech.SetOutputToDefaultAudioDevice();
            speech.SpeakAsync(text);
        }

        private async void Init(object sender, RoutedEventArgs e)
        {
            
            this.nativeWindow = new WindowInteropHelper(this);
            HwndSource hwndSource = HwndSource.FromHwnd(this.nativeWindow.Handle);
            this.selectionReader = new SelectionReader(this.nativeWindow.Handle);
            hwndSource?.AddHook(this.selectionReader.WindowProc);

            if (this.environmentWindow != IntPtr.Zero)
            {
                await this.SaySelection(this.environmentWindow);
                this.Close();
            }
        }
    }
}
