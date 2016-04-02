namespace Hackathon.CodeSharp
{
    using System;
    using System.Diagnostics;
    using System.Drawing;
    using System.Globalization;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;
    using EnvDTE;
    using EnvDTE80;
    using Microsoft.VisualStudio.ComponentModelHost;
    using Microsoft.VisualStudio.Shell;
    using Microsoft.VisualStudio.TestWindow.Extensibility;
    using Task = System.Threading.Tasks.Task;
    using System.Speech.Synthesis;
    using Microsoft.VisualStudio.PlatformUI;
    using Microsoft.VisualStudio.Shell.Interop;
    using CamCaptureLib;
    using System.Speech.Recognition;
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.1", IconResourceID = 400)]
    [Guid(GuidList.codeSharpPkgString)]
    [ProvideAutoLoad("{f1536ef8-92ec-443c-9ed7-fdadf150da82}")]
    public sealed class CodeSharpPackage : Package, IDisposable
    {
        private DTE2 _dte;
        private BuildEvents buildEvents;
        private DebuggerEvents debugEvents;
        private OptionsDialog _options = null;
        private CommandTable _cache;
        private SpeechRecognitionEngine _rec;
        private bool _isEnabled, _isListening;
        private string _rejected;
        private const float _minConfidence = 0.80F; // A value between 0 and 1
        int buildFailedCount = 0; 
        //private Players players = null;

        public CodeSharpPackage()
        {
            Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering constructor for: {0}", ToString()));
        }

        #region Package Members

        protected override void Initialize()
        {
            Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering Initialize() of: {0}", ToString()));
            base.Initialize();

            _dte = (DTE2)GetService(typeof(DTE));
            //_cache = new CommandTable(_dte, this);

            buildEvents = _dte.Events.BuildEvents;
            debugEvents = _dte.Events.DebuggerEvents;

            SetupEventHandlers();
            //InitializeSpeechRecognition();
        }

        public void Talk(string text)
        {
            SpeechSynthesizer synth = new SpeechSynthesizer();

            // Configure the audio output. 
            synth.SetOutputToDefaultAudioDevice();

            // Speak a string.
            synth.Speak(text);

        }

        private void SetupEventHandlers()
        {
            buildEvents.OnBuildDone += (scope, action) =>
            {
                if (Options.IsBeepOnBuildComplete)
                {
                    HandleEventSafe(EventType.BuildCompleted, "Build has been completed.");
                }
            };

            buildEvents.OnBuildProjConfigDone += OnProjectBuildFinished;
            
            debugEvents.OnExceptionThrown += DebugEvents_OnExceptionThrown;
            debugEvents.OnExceptionNotHandled += DebugEvents_OnExceptionThrown;

            debugEvents.OnEnterBreakMode += delegate(dbgEventReason reason, ref dbgExecutionAction action)
            {
                if (reason != dbgEventReason.dbgEventReasonStep && Options.IsBeepOnBreakpointHit)
                {
                    HandleEventSafe(EventType.BreakpointHit, "Breakpoint was hit.");
                }
            };

            var componentModel =
                GetGlobalService(typeof(SComponentModel)) as IComponentModel;

            if (componentModel == null)
            {
                Debug.WriteLine("componentModel is null");
                return;
            }

            var operationState = componentModel.GetService<IOperationState>();
            operationState.StateChanged += OperationStateOnStateChanged;
        }

        private void DebugEvents_OnExceptionThrown(string ExceptionType, string Name, int Code, string Description, ref dbgExceptionAction ExceptionAction)
        {
            this.Talk("Here is an exception by " + Name);
        }

        private async void OnProjectBuildFinished(string Project, string ProjectConfig, string Platform, string SolutionConfig, bool Success)
        {
            if (Success) return;
            // Speak a string.
            //try
            //{
            //    CamCam cam = new CamCam();
            //    var result = await cam.GetEmotions();
            //    this.Talk("build error");
            //}
            //catch (Exception e) { this.Talk(e.Message); }

            this.Talk("you break the build and it is good to fail sometimes");
            
        }

        private OptionsDialog Options
        {
            get
            {
                if (_options == null)
                {
                    _options = (OptionsDialog)GetDialogPage(typeof(OptionsDialog));
                }
                return _options;
            }
        }

      

        private void HandleEventSafe(EventType eventType, string messageText, ToolTipIcon icon = ToolTipIcon.Info)
        {
            if(eventType == EventType.BreakpointHit)
                 Talk("breakpoint hit");
            if (!ShouldPerformNotificationAction())
            {
                return;
            }

     
            ShowNotifyMessage(messageText, icon);

        

        }

        private void ShowNotifyMessage(string messageText, ToolTipIcon icon = ToolTipIcon.Info)
        {
            if (!_options.ShowTrayNotifications)
            {
                return;
            }

            if (Options.ShowTrayDisableMessage)
            {
                string autoAppendMessage = Environment.NewLine + "You can disable this notification in:" + Environment.NewLine + "Tools->Options->Ding->Show tray notifications";
                messageText = string.Format("{0}{1}", messageText, autoAppendMessage);
            }

            Task.Run(async () =>
                {
                    var tray = new NotifyIcon
                    {
                        Icon = SystemIcons.Application,
                        BalloonTipIcon = icon,
                        BalloonTipText = messageText,
                        BalloonTipTitle = "Visual Studio Ding extension",
                        Visible = true
                    };

                    tray.ShowBalloonTip(5000);
                    await Task.Delay(5000);
                    tray.Icon = (Icon)null;
                    tray.Visible = false;
                    tray.Dispose();
                });
        }

        private bool ShouldPerformNotificationAction()
        {
            if (!Options.IsBeepOnlyWhenVisualStudioIsInBackground)
            {
                return true;
            }
            return Options.IsBeepOnlyWhenVisualStudioIsInBackground && !WinApiHelper.ApplicationIsActivated();
        }

        private void OperationStateOnStateChanged(object sender, OperationStateChangedEventArgs operationStateChangedEventArgs)
        {
           // if (Options.IsBeepOnTestComplete && operationStateChangedEventArgs.State.HasFlag(TestOperationStates.TestExecutionFinished))
            {
                try
                {
                    // Issue #8: VS 2015 stops working when looking at Test Manager Window #8 
                    // This extention can't take dependency on Microsoft.VisualStudio.TestWindow.Core.dll
                    // Because it will crash VS 2015. But DominantTestState is defined in that assembly.
                    // So as a workaround - cast it to dynamic (ewww, but alternative - to create new project/build and publish it separately.)
                    var testOperation = (dynamic)(operationStateChangedEventArgs.Operation);
                    var dominantTestState = (TestState)testOperation.DominantTestState;
                    var isTestsFailed = dominantTestState == TestState.Failed;
                    var eventType = isTestsFailed ? EventType.TestsCompletedFailure : EventType.TestsCompletedSuccess;
                    if (isTestsFailed)
                    {
                        Talk("Test execution is failed!!");
                        HandleEventSafe(eventType, "Test execution failed!", ToolTipIcon.Error);
                    }
                    else
                    {
                        HandleEventSafe(eventType, "Test execution has been completed.");
                    }
                }
                catch (Exception ex)
                {
                    ActivityLog.LogError(GetType().FullName, ex.Message);
                    // Unable to get dominate test status, beep default sound for test
                    HandleEventSafe(EventType.TestsCompletedSuccess, "Test execution has been completed.");
                }
            }
        }
        #endregion

        public void Dispose()
        {
            //players.Dispose();
        }



        private void InitializeSpeechRecognition()
        {
            try
            {
                var arr = new string[_cache.Commands.Keys.Count];
                _cache.Commands.Keys.CopyTo(arr, 0);
                var c = new Choices(arr);
                var gb = new GrammarBuilder(c);
                var g = new Grammar(gb);


                _rec = new SpeechRecognitionEngine();
                _rec.LoadGrammar(g);
                _rec.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(sRecognize_SpeechRecognized);
                _rec.SetInputToDefaultAudioDevice();
                _rec.RecognizeAsync(RecognizeMode.Multiple);

                
                _isEnabled = true;
            }
            catch (Exception ex)
            {
                Talk(ex.Message);
                //Logger.Log(ex);
            }
        }

        void sRecognize_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
         //   MessageBox.Show("speech is:" + e.Result.Text.ToString());

            try
            {
              

                _rec.RecognizeAsyncStop();
                _isListening = false;

               if (e.Result != null && e.Result.Confidence > _minConfidence)
                { // Speech matches a command
                    _cache.ExecuteCommand(e.Result.Text);
                    MessageBox.Show("speech is: confident  " + e.Result.Text.ToString());
                    // var props = new Dictionary<string, string> { { "phrase", e.Result.Text } };
                    //Telemetry.TrackEvent("Match", props);
                }
                else if (string.IsNullOrEmpty(_rejected))
                { // Speech didn't match a command
                    _dte.StatusBar.Text = "I didn't quite get that. Please try again.";
                    // Telemetry.TrackEvent("No match");
                }
         
                else
                { // No match or timeout
                    _dte.StatusBar.Clear();
                }
            }
            catch (Exception ex)
            {
                _dte.StatusBar.Clear();
                // Logger.Log(ex);
            }
            //if (e.Result.Confidence >= 0.3)
            // Talk("speech is:" + e.Result.Text.ToString());
        }

        private void OnListening(object sender, EventArgs e)
        {
            try
            {
                if (!_isEnabled)
                {
                    SetupVoiceRecognition();
                }
                else if (!_isListening)
                {
                    _isListening = true;
                    _rec.RecognizeAsync();
                    _dte.StatusBar.Text = "I'm listening...";
                }
            }
            catch (Exception ex)
            {
                Talk(ex.Message);
                //Logger.Log(ex);
            }
        }

        private void OnSpeechHypothesized(object sender, SpeechHypothesizedEventArgs e)
        {
            if (string.IsNullOrEmpty(_rejected))
                _dte.StatusBar.Text = "I'm listening... (" + e.Result.Text + " " + Math.Round(e.Result.Confidence * 100) + "%)";
            Talk("I am listening");
        }

        private void OnSpeechRecognitionRejected(object sender, SpeechRecognitionRejectedEventArgs e)
        {
            if (e.Result.Text != "yes" && e.Result.Confidence > 0.5F)
            {
                Talk("I am listening");

                _rejected = e.Result.Text;
                _dte.StatusBar.Text = "Did you mean " + e.Result.Text + "? (say yes or no)";
            }
        }

        private void OnSpeechRecognized(object sender, RecognizeCompletedEventArgs e)
        {
            try
            {
               

                //_rec.RecognizeAsyncStop();
                _isListening = false;

                
                 if (e.Result != null && e.Result.Confidence > _minConfidence)
                { // Speech matches a command
                    MessageBox.Show(e.Result.Text);
                    _cache.ExecuteCommand(e.Result.Text);
                   // var props = new Dictionary<string, string> { { "phrase", e.Result.Text } };
                    //Telemetry.TrackEvent("Match", props);
                }
                else if (string.IsNullOrEmpty(_rejected))
                { // Speech didn't match a command
                   // _dte.StatusBar.Text = "I didn't quite get that. Please try again.";
                   // Telemetry.TrackEvent("No match");
                }
                else if (e.Result == null && !string.IsNullOrEmpty(_rejected) && !e.InitialSilenceTimeout)
                { // Keep listening when asked about rejected speech
                   // _rec.RecognizeAsync();
                   // Telemetry.TrackEvent("Low confidence");
                }
                else
                { // No match or timeout
                   // _dte.StatusBar.Clear();
                }
            }
            catch (Exception ex)
            {
                _dte.StatusBar.Clear();
               // Logger.Log(ex);
            }
        }

        private static void SetupVoiceRecognition()
        {
            string message = "Do you want to learn how to setup voice recognition in Windows?";
            var answer = MessageBox.Show(message, "Code Sharp", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (answer == DialogResult.Yes)
            {
                string url = "http://windows.microsoft.com/en-US/windows-8/using-speech-recognition/";
                System.Diagnostics.Process.Start(url);
            }
        }
    }
}
