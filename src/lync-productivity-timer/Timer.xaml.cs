using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;
using Microsoft.Lync.Model;
using System.Media;

namespace lync_productivity_timer
{
    /// <summary>
    /// Interaction logic for Timer.xaml
    /// </summary>
    public partial class Timer : Window
    {
        private DispatcherTimer updateTimer;
        private short minutes;
        private FlashWindowHelper blinker;

        public Timer()
        {
            InitializeComponent();
            blinker = new FlashWindowHelper(Application.Current);
            updateTimer = new DispatcherTimer();
            updateTimer.Interval = new TimeSpan(0, 1, 0);
            updateTimer.Tick += new EventHandler(updateTimer_Tick);
            updateTimer.IsEnabled = false;

            txtAwayMessageTemplate.Text = "In 'flow' for the next {0}. Please wait until {1} or email me if it's not urgent. Thanks!";
            lblReplacements.Content = "Replacements:" + Environment.NewLine + "  {0}: 'xx minute(s)' - countdown" + Environment.NewLine + "  {1}: 'xx:xx' - end time";
        }

        private void TriggerAlarm()
        {
            // TODO: Play custom sound?
            SystemSounds.Asterisk.Play();

            // Blink the window
            blinker.FlashApplicationWindow();

            // TODO: Replace with a tray notification?
            MessageBox.Show("Productivity session has ended");
        }

        void updateTimer_Tick(object sender, EventArgs e)
        {
            minutes -= 1;
            if (minutes <= 0)
            {
                // Alarm!
                TriggerAlarm();
                btnStop_Click(this, null);
            }
            else
            {
                string remainaing = String.Format("{0} minute{1}", minutes, minutes == 1 ? "" : "s");
                string endTime = DateTime.Now.AddMinutes(minutes).ToShortTimeString();

                // Update countdown label                
                lblTimeRemaining.Content = String.Format("{0} ({1})", remainaing, endTime);
                // Update status
                UpdateLync(PublishableContactInformationType.PersonalNote, String.Format(txtAwayMessageTemplate.Text, remainaing, endTime));
            }
        }

        private void btnStart_Click(object sender, RoutedEventArgs e)
        {
            btnStart.IsEnabled = false;
            
            minutes = Int16.Parse(txtLength.Text);
            updateTimer.IsEnabled = true;
            updateTimer.Start();
            
            UpdateLync(PublishableContactInformationType.Availability, ContactAvailability.DoNotDisturb);    

            minutes++;
            updateTimer_Tick(this, null);

            btnStop.IsEnabled = true;
        }

        private void btnStop_Click(object sender, RoutedEventArgs e)
        {
            btnStop.IsEnabled = false;

            updateTimer.Stop();
            updateTimer.IsEnabled = false;

            UpdateLync(PublishableContactInformationType.Availability, ContactAvailability.Free);
            UpdateLync(PublishableContactInformationType.PersonalNote, "");

            lblTimeRemaining.Content = "0 minutes";
           
            btnStart.IsEnabled = true;
        }

        private void UpdateLync(PublishableContactInformationType type, object value)
        {
            var client = LyncClient.GetClient();
            var contactInfo = new Dictionary<PublishableContactInformationType, object>();
            contactInfo.Add(type, value);
            try
            {

                var publish = client.Self.BeginPublishContactInformation(contactInfo, null, null);
                client.Self.EndPublishContactInformation(publish);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ExitApp(object sender, EventArgs e)
        {
            if (updateTimer.IsEnabled)
            {
                if (MessageBox.Show("You have a timer running. Exit anyway?", "Running Timer", MessageBoxButton.YesNo) != MessageBoxResult.Yes)
                {
                    if (e is System.ComponentModel.CancelEventArgs)
                    {
                        ((System.ComponentModel.CancelEventArgs) e).Cancel = true;
                    }
                    return;
                }

                btnStop_Click(this, null);
            }

            Application.Current.Shutdown();
        }

        private void MenuItem_About_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("(c) 2014 Colin O'Dell", "About " + this.Title, MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void Window_Activated(object sender, EventArgs e)
        {
            blinker.StopFlashing();
        }
    }
}
