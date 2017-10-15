namespace Microsoft.Samples.Kinect.DiscreteGestureBasics
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.ComponentModel;
    using System.Windows;
    using System.Windows.Controls;
    using Microsoft.Kinect;
    using Microsoft.Kinect.VisualGestureBuilder;
    using System.Linq;
    using System.IO;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Windows.Documents;
    using System.Windows.Media;
    using Microsoft.Speech.AudioFormat;
    using Microsoft.Speech.Recognition;
    using System.IO.Ports;
    using System.Speech.Synthesis;


    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1001:TypesThatOwnDisposableFieldsShouldBeDisposable",
       Justification = "In a full-fledged application, the SpeechRecognitionEngine object should be properly disposed. For the sake of simplicity, we're omitting that code in this sample.")]

    /// <summary>
    /// Interaction logic for the MainWindow
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        public static SerialPort sp = new SerialPort();
        System.Speech.Synthesis.SpeechSynthesizer speaker = new System.Speech.Synthesis.SpeechSynthesizer();
        public static Boolean spflagl = false, spflagr = false;

        /// <summary> Active Kinect sensor </summary>
        private KinectSensor kinectSensor = null;
        private KinectAudioStream convertStream = null;
        private SpeechRecognitionEngine speechEngine = null;
        /// <summary> Array for the bodies (Kinect will track up to 6 people simultaneously) </summary>
        private Body[] bodies = null;

        /// <summary> Reader for body frames </summary>
        private BodyFrameReader bodyFrameReader = null;

        /// <summary> Current status text to display </summary>
        private string statusText = null;

        /// <summary> KinectBodyView object which handles drawing the Kinect bodies to a View box in the UI </summary>
        private KinectBodyView kinectBodyView = null;
        
        /// <summary> List of gesture detectors, there will be one detector created for each potential body (max of 6) </summary>
        private List<GestureDetector> gestureDetectorList = null;

        /// <summary>
        /// Initializes a new instance of the MainWindow class
        /// </summary>
        /// 
      


        public MainWindow()
        {
            // only one sensor is currently supported
            this.kinectSensor = KinectSensor.GetDefault();
            
            // set IsAvailableChanged event notifier
            this.kinectSensor.IsAvailableChanged += this.Sensor_IsAvailableChanged;

            // open the sensor
            this.kinectSensor.Open();

            // set the status text
            this.StatusText = this.kinectSensor.IsAvailable ? Properties.Resources.RunningStatusText
                                                            : Properties.Resources.NoSensorStatusText;

            // open the reader for the body frames
            this.bodyFrameReader = this.kinectSensor.BodyFrameSource.OpenReader();

            // set the BodyFramedArrived event notifier
            this.bodyFrameReader.FrameArrived += this.Reader_BodyFrameArrived;

            // initialize the BodyViewer object for displaying tracked bodies in the UI
            this.kinectBodyView = new KinectBodyView(this.kinectSensor);

            // initialize the gesture detection objects for our gestures
            this.gestureDetectorList = new List<GestureDetector>();

            // initialize the MainWindow
            this.InitializeComponent();
            try
            {
                sp.PortName = "COM6";
                sp.BaudRate = 9600;
                sp.Open();
            }

            catch (Exception)
            {
                MessageBox.Show("Please give a valid port number or check your connection");
            }

            
            speaker.SelectVoiceByHints(System.Speech.Synthesis.VoiceGender.Female);

            
            // set our data context objects for display in UI
            this.DataContext = this;
            this.kinectBodyViewbox.DataContext = this.kinectBodyView;

            // create a gesture detector for each body (6 bodies => 6 detectors) and create content controls to display results in the UI
           // int col0Row = 0;
          //  int col1Row = 0;
            int maxBodies = this.kinectSensor.BodyFrameSource.BodyCount;
            //int maxBodies = 1;
            for (int i = 0; i < maxBodies; ++i)
            {
                GestureResultView result = new GestureResultView(i, false, false, 0.0f);
                GestureDetector detector = new GestureDetector(this.kinectSensor, result);
                this.gestureDetectorList.Add(detector);                
                
                // split gesture results across the first two columns of the content grid
                ContentControl contentControl = new ContentControl();
                contentControl.Content = this.gestureDetectorList[i].GestureResultView;

                if (this.kinectSensor != null)
                {
                    // open the sensor
                    this.kinectSensor.Open();


                    // grab the audio stream
                    IReadOnlyList<AudioBeam> audioBeamList = this.kinectSensor.AudioSource.AudioBeams;
                    System.IO.Stream audioStream = audioBeamList[0].OpenInputStream();

                    // create the convert stream
                    this.convertStream = new KinectAudioStream(audioStream);
                }
                else
                {
                    // on failure, set the status text

                    return;
                }

                RecognizerInfo ri = TryGetKinectRecognizer();

                if (null != ri)
                {
                    this.speechEngine = new SpeechRecognitionEngine(ri.Id);

               

                    // Create a grammar from grammar definition XML file.
                    using (var memoryStream = new MemoryStream(Encoding.ASCII.GetBytes(Properties.Resources.SpeechGrammar)))
                    {
                        var g = new Grammar(memoryStream);
                        this.speechEngine.LoadGrammar(g);
                    }

                    this.speechEngine.SpeechRecognized += this.SpeechRecognized;
                    this.speechEngine.SpeechRecognitionRejected += this.SpeechRejected;

                    // let the convertStream know speech is going active
                    this.convertStream.SpeechActive = true;

                    // For long recognition sessions (a few hours or more), it may be beneficial to turn off adaptation of the acoustic model. 
                    // This will prevent recognition accuracy from degrading over time.
                    ////speechEngine.UpdateRecognizerSetting("AdaptationOn", 0);

                    this.speechEngine.SetInputToAudioStream(
                        this.convertStream, new SpeechAudioFormatInfo(EncodingFormat.Pcm, 16000, 16, 1, 32000, 2, null));
                    this.speechEngine.RecognizeAsync(RecognizeMode.Multiple);
                }
                else
                {
                    //this.statusBarText.Text = Properties.Resources.NoSpeechRecognizer;
                }
            }

            ContentControl contentControl2 = new ContentControl();
            contentControl2.Content = this.gestureDetectorList[0].GestureResultView;
            this.contentGrid.Children.Add(contentControl2);
        }

        private static Speech.Recognition.RecognizerInfo TryGetKinectRecognizer()
        {
            IEnumerable<Speech.Recognition.RecognizerInfo> recognizers;

            // This is required to catch the case when an expected recognizer is not installed.
            // By default - the x86 Speech Runtime is always expected. 
            try
            {
                recognizers = SpeechRecognitionEngine.InstalledRecognizers();
            }
            catch (COMException)
            {
                return null;
            }

            foreach (Speech.Recognition.RecognizerInfo recognizer in recognizers)
            {
                string value;
                recognizer.AdditionalInfo.TryGetValue("Kinect", out value);
                if ("True".Equals(value, StringComparison.OrdinalIgnoreCase) && "en-US".Equals(recognizer.Culture.Name, StringComparison.OrdinalIgnoreCase))
                {
                    return recognizer;
                }
            }

            return null;
        }

        /// <summary>
        /// INotifyPropertyChangedPropertyChanged event to allow window controls to bind to changeable data
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// Gets or sets the current status text to display
        /// </summary>
        public string StatusText
        {
            get
            {
                return this.statusText;
            }

            set
            {
                if (this.statusText != value)
                {
                    this.statusText = value;

                    // notify any bound elements that the text has changed
                    if (this.PropertyChanged != null)
                    {
                        this.PropertyChanged(this, new PropertyChangedEventArgs("StatusText"));
                    }
                }
            }
        }

        /// <summary>
        /// Execute shutdown tasks
        /// </summary>
        /// <param name="sender">object sending the event</param>
        /// <param name="e">event arguments</param>
        private void MainWindow_Closing(object sender, CancelEventArgs e)
        {
            try
            {
                sp.Close();
            }
            catch (Exception)
            {

                MessageBox.Show("First Connect and then disconnect");
            }

            if (null != this.convertStream)
            {
                this.convertStream.SpeechActive = false;
            }

            if (null != this.speechEngine)
            {
                this.speechEngine.SpeechRecognized -= this.SpeechRecognized;
                this.speechEngine.SpeechRecognitionRejected -= this.SpeechRejected;
                this.speechEngine.RecognizeAsyncStop();
            }

            if (this.bodyFrameReader != null)
            {
                // BodyFrameReader is IDisposable
                this.bodyFrameReader.FrameArrived -= this.Reader_BodyFrameArrived;
                this.bodyFrameReader.Dispose();
                this.bodyFrameReader = null;
            }

            if (this.gestureDetectorList != null)
            {
                // The GestureDetector contains disposable members (VisualGestureBuilderFrameSource and VisualGestureBuilderFrameReader)
                foreach (GestureDetector detector in this.gestureDetectorList)
                {
                    detector.Dispose();
                }

                this.gestureDetectorList.Clear();
                this.gestureDetectorList = null;
            }

            if (this.kinectSensor != null)
            {
                this.kinectSensor.IsAvailableChanged -= this.Sensor_IsAvailableChanged;
                this.kinectSensor.Close();
                this.kinectSensor = null;
            }
        }

        /// <summary>
        /// Handles the event when the sensor becomes unavailable (e.g. paused, closed, unplugged).
        /// </summary>
        /// <param name="sender">object sending the event</param>
        /// <param name="e">event arguments</param>
        private void Sensor_IsAvailableChanged(object sender, IsAvailableChangedEventArgs e)
        {
            // on failure, set the status text
            this.StatusText = this.kinectSensor.IsAvailable ? Properties.Resources.RunningStatusText
                                                            : Properties.Resources.SensorNotAvailableStatusText;
        }

        /// <summary>
        /// Handles the body frame data arriving from the sensor and updates the associated gesture detector object for each body
        /// </summary>
        /// <param name="sender">object sending the event</param>
        /// <param name="e">event arguments</param>
        private void Reader_BodyFrameArrived(object sender, BodyFrameArrivedEventArgs e)
        {
            bool dataReceived = false;

            using (BodyFrame bodyFrame = e.FrameReference.AcquireFrame())
            {
                if (bodyFrame != null)
                {
                    if (this.bodies == null)
                    {
                        // creates an array of 6 bodies, which is the max number of bodies that Kinect can track simultaneously
                        this.bodies = new Body[bodyFrame.BodyCount];
                    }

                    // The first time GetAndRefreshBodyData is called, Kinect will allocate each Body in the array.
                    // As long as those body objects are not disposed and not set to null in the array,
                    // those body objects will be re-used.
                    bodyFrame.GetAndRefreshBodyData(this.bodies);
                    dataReceived = true;
                }
            }

            if (dataReceived)
            {
                // visualize the new body data
                this.kinectBodyView.UpdateBodyFrame(this.bodies);

                // we may have lost/acquired bodies, so update the corresponding gesture detectors
                if (this.bodies != null)
                {
                    Body body = this.bodies[0];
                    if (this.bodies.Where(b => b.IsTracked == true).Count() != 0)
                        body = this.bodies.Where(b => b.IsTracked == true).First();
                    if(body != null)
                    {
                        ulong trackingId = body.TrackingId;
                       if (trackingId != this.gestureDetectorList[0].TrackingId)
                        {
                            this.gestureDetectorList[0].TrackingId = trackingId;
                            this.gestureDetectorList[0].IsPaused = trackingId == 0;
                        }
                    }
                }
            }
        }
        private void SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            // Speech utterance confidence below which we treat speech as if it hadn't been heard
            Console.WriteLine("Inside speechrecognized event");
            const double ConfidenceThreshold = 0.3;

            if (e.Result.Confidence >= ConfidenceThreshold)
            {
                Console.WriteLine(e.Result.Confidence);
                switch (e.Result.Semantics.Value.ToString())
                {
                    case "FORWARD":
                        sp.WriteLine("256");
                        sp.WriteLine("170");
                        Console.WriteLine(e.Result.Text);
                        Console.WriteLine("*************LEFT**************");
                        if (spflagl == true)
                        {    
                            speaker.SpeakAsync("I'm sorry, it is on");
                        }
                        else
                        {
                            // speaker.SpeakAsync("I have switched on the left L E D");
                            speaker.SpeakAsync("Done");
                            spflagl = true;
                        }
                        break;

                    case "BACKWARD":
                        sp.WriteLine("257");
                        sp.WriteLine("170");
                        Console.WriteLine("*************RIGHT**************");
                        if (spflagr == true)
                        {
                            //speaker.SpeakAsync("I'm sorry, it is on");
                            speaker.SpeakAsync("Done");
                        }
                        else
                        {
                            //speaker.SpeakAsync("I have switched on the right L E D");
                            speaker.SpeakAsync("Done");
                            spflagr = true;
                        }
                        break;

                    case "LEFT":
                        sp.WriteLine("500");
                       
                        if (spflagl == true || spflagr == true)
                        {
                            spflagl = false;
                            spflagr = false;
                           // speaker.SpeakAsync("Switching off the LEDs");
                            speaker.SpeakAsync("Done");
                        }
                        break;

                    case "RIGHTLOW":
                            sp.WriteLine("257");
                            sp.WriteLine("84");
                        spflagr = true;
                            //speaker.SpeakAsync("Decreased brightness of the right L E D");
                        speaker.SpeakAsync("Done");
                        break;

                    case "RIGHTHIGH":
                        sp.WriteLine("257");
                        sp.WriteLine("255");
                        spflagr = true;
                        //    speaker.SpeakAsync("increased brightness of the right L E D");
                        speaker.SpeakAsync("Done");
                        break;
                    case "LEFTLOW":
                        sp.WriteLine("256");
                        sp.WriteLine("84");
                        spflagl = true;
                      //  speaker.SpeakAsync("Decreased brightness of the left L E D");
                        speaker.SpeakAsync("Done");
                        break;
                    case "LEFTHIGH":
                        sp.WriteLine("256");
                        sp.WriteLine("255");
                        spflagl = true;
                        speaker.SpeakAsync("Done");
                        //speaker.SpeakAsync("Increased brightness of the left L E D");
                        break;
                    case "BOTH_ON":
                        Console.Write("Both on");
                        Console.WriteLine(e.Result.Text);
                        if (spflagl == true && spflagr == true)
                            speaker.SpeakAsync("Already on");
                        {
                            sp.WriteLine("256");
                            sp.WriteLine("255");
                            sp.WriteLine("257");
                            sp.WriteLine("255");
                            speaker.SpeakAsync("Done");
                        }


                        break;
                }
            }
        }

        private void SpeechRejected(object sender, SpeechRecognitionRejectedEventArgs e)
        {
            return;
        }
    }
}
