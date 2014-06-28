using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace GmailNotifier
{
    public partial class MainForm : Form
    {
        #region Members
        private NotifyIcon _notifyIcon;
        private ContextMenu _contextMenu;
        private Timer _timer;
        private GmailConnection _connection;
        #endregion

        #region Initialization
        /// <summary>
        /// Constructor.
        /// </summary>
        public MainForm()
        {
            InitializeComponent();
            InitializeMenu();
            InitializeIcon();

            Load += MainForm_Load;
            FormClosing += MainForm_FormClosing;
            FormClosed += MainForm_FormClosed;
        }

        /// <summary>
        /// Connects to the Gmail API and starts the timer.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void MainForm_Load(object sender, EventArgs e)
        {
            ShowInTaskbar = false;
            Visible = false;
            WindowState = FormWindowState.Minimized;

            _connection = new GmailConnection("me");
            _connection.Connect();

            _timer = new Timer();
            _timer.Tick += Timer_Tick;
            _timer.Interval = 5 * 60 * 1000;
            _timer.Enabled = true;
            _timer.Start();

            // Check mail immediately
            Timer_Tick(this, new EventArgs());
        }

        /// <summary>
        /// Initializes the context menu for the icon.
        /// </summary>
        private void InitializeMenu()
        {
            _contextMenu = new ContextMenu();

            _contextMenu.MenuItems.AddRange(
                new MenuItem[]
                {
                    GetMenuItem("Check Now", Timer_Tick),
                    GetMenuItem("Sign Out", SignOut_Click),
                    GetMenuItem("-", null),
                    GetMenuItem("Exit", Exit_Click)
                }
            );
        }

        /// <summary>
        /// Initializes the taskbar icon.
        /// </summary>
        private void InitializeIcon()
        {
            _notifyIcon = new NotifyIcon(components);

            _notifyIcon.Icon = Properties.Resources.gmail_off;
            _notifyIcon.ContextMenu = _contextMenu;
            _notifyIcon.Text = Properties.Resources.ApplicationName;
            _notifyIcon.Visible = true;
            _notifyIcon.DoubleClick += NotifyIcon_DoubleClick;
        }
        #endregion

        #region Methods
        /// <summary>
        /// Gets a new menu item for the context menu.
        /// </summary>
        /// <param name="text"></param>
        /// <param name="eventHandler"></param>
        /// <returns></returns>
        private MenuItem GetMenuItem(string text, EventHandler eventHandler)
        {
            MenuItem item = new MenuItem(text);

            if (eventHandler != null) item.Click += eventHandler;

            return item;
        }

        /// <summary>
        /// Disables the notification icon.
        /// </summary>
        private void DisableIcon()
        {
            if (_notifyIcon != null)
            {
                _notifyIcon.Visible = false;
                _notifyIcon.Icon = null;
                _notifyIcon.Visible = false;
                _notifyIcon.Dispose();
                _notifyIcon = null;
            }
        }
        #endregion

        #region Event Handlers
        /// <summary>
        /// Checks for new emails.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void Timer_Tick(object sender, EventArgs e)
        {
            if (_connection.IsConnected)
            {
                bool hasEmails = _connection.CheckMessages();

                if (hasEmails)
                {
                    _notifyIcon.Icon = Properties.Resources.gmail_on;
                }
                else
                {
                    _notifyIcon.Icon = Properties.Resources.gmail_off;
                }

                _notifyIcon.Text = string.Format("{0} unread email{1}", _connection.MailCount, _connection.MailCount > 1 ? "s" : "");
            }
        }

        /// <summary>
        /// Opens up Gmail in a browser.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void NotifyIcon_DoubleClick(object sender, EventArgs e)
        {
            Process.Start(@"https://mail.google.com");
        }

        /// <summary>
        /// Disconnects from Gmail.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void SignOut_Click(object sender, EventArgs e)
        {
            if (_connection != null && _connection.IsConnected)
            {
                _connection.Disconnect();

                Application.Exit();
            }
        }

        /// <summary>
        /// Hides the form instead of closing it when the control box is used.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                WindowState = FormWindowState.Minimized;
                Visible = false;
                e.Cancel = true;
            }
        }

        /// <summary>
        /// Exits the application when the exit menu command is called.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void Exit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        /// <summary>
        /// Removes the icon when the application exits.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            DisableIcon();
        }
        #endregion
    }
}
