using ReadSelectedText;
using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;
using System.Speech.Synthesis;
using System.Threading;

namespace ReadSelectedTextApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private WindowInteropHelper nativeWindow;
        private SelectionReader selectionReader;
        private readonly IntPtr environmentWindow;
        private readonly SpeechSynthesizer speech = new SpeechSynthesizer();

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
            if (this.speech.State != SynthesizerState.Ready)
            {
                this.speech.SpeakAsyncCancelAll();
            }
            else
            {
                string text = await this.selectionReader.GetSelectedText(activeWindow);
                this.speech.SetOutputToDefaultAudioDevice();
                this.speech.SpeakAsync(text);
            }
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
