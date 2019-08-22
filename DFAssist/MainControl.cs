using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Speech.Synthesis;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml;
using Windows.UI.Notifications;
using Advanced_Combat_Tracker;
using DFAssist.DataModel;
using DFAssist.Shell;
using Microsoft.Win32;
using Timer = System.Windows.Forms.Timer;
using Telegram.Bot;
using PushbulletSharp;

namespace DFAssist
{
    public class MainControl : UserControl, IActPluginV1
    {
        private const string AppId = "Advanced Combat Tracker";

        private readonly ConcurrentDictionary<int, ProcessNet> _networks;
        private readonly string _settingsFile;
        private readonly SpeechSynthesizer _synth;

        private TabControl _appTabControl;
        private Label _appTitle;
        private Button _button1;
        private LinkLabel _copyrightLink;
        private CheckBox _disableToasts;
        private CheckBox _enableLegacyToast;
        private CheckBox _enableTestEnvironment;
        private GroupBox _generalSettings;
        private LegacyToast _lastToast;

        private bool _isPluginEnabled;
        private Label _label1;
        private Label _labelStatus;
        private TabPage _labelTab;
        private ComboBox _languageComboBox;
        private TextBox _languageValue;
        private bool _mainFormIsLoaded;
        private TableLayoutPanel _mainTableLayout;
        private TabPage _mainTabPage;
        private CheckBox _persistToasts;
        private bool _pluginInitializing;
        private RichTextBox _richTextBox1;
        private Language _selectedLanguage;
        private TabPage _settingsPage;
        private TableLayoutPanel _settingsTableLayout;
        private GroupBox _testSettings;
        private Timer _timer;
        private GroupBox _toastSettings;
        private CheckBox _ttsCheckBox;
        private GroupBox _ttsSettings;
        private SettingsSerializer _xmlSettingsSerializer;
        private GroupBox _telegramSettings;
        private TextBox _chatIdTextBox;
        private Label _ChatIdLabel;
        private TextBox _telegramTokenTextBox;
        private Label _tokenLabel;
        private CheckBox _telegramCheckBox;
        private GroupBox _pushbulletSettings;
        private TextBox _pushbulletDeviceIdTextBox;
        private Label _pushbulletDeviceIdlabel;
        private TextBox _pushbulletTokenTextBox;
        private Label _pushbulletTokenLabel;
        private CheckBox _pushbulletCheckbox;
        private Panel _settingsPanel;

        #region Load Methods

        private void LoadData(Language defaultLanguage = null)
        {
            var newLanguage = defaultLanguage ?? (Language)_languageComboBox.SelectedItem;
            if(_selectedLanguage != null && newLanguage.Code.Equals(_selectedLanguage.Code))
                return;

            _selectedLanguage = newLanguage;
            _languageValue.Text = _selectedLanguage.Name;
            Localization.Initialize(_selectedLanguage.Code);
            Data.Initialize(_selectedLanguage.Code);

            UpdateTranslations();
        }

        #endregion

        #region WinForm Required

        public MainControl()
        {
            InitializeComponent();

            _synth = new SpeechSynthesizer();

            AppDomain.CurrentDomain.AssemblyResolve += CurrentDomainOnAssemblyResolve;

            _settingsFile = Path.Combine(ActGlobals.oFormActMain.AppDataFolder.FullName, "Config", "DFAssist.config.xml");
            _networks = new ConcurrentDictionary<int, ProcessNet>();

            foreach(Form formLoaded in Application.OpenForms)
            {
                if(formLoaded != ActGlobals.oFormActMain)
                    continue;

                _mainFormIsLoaded = true;
                break;
            }
        }

        private Assembly CurrentDomainOnAssemblyResolve(object sender, ResolveEventArgs e)
        {
            // if any of the assembly cannot be loaded, then the plugin cannot be started
            if(!AssemblyResolver.LoadAssembly(e, _labelStatus, out var result))
                throw new Exception("Assembly load failed.");

            return result;
        }

        private void InitializeComponent()
        {
            this._label1 = new System.Windows.Forms.Label();
            this._languageValue = new System.Windows.Forms.TextBox();
            this._languageComboBox = new System.Windows.Forms.ComboBox();
            this._enableTestEnvironment = new System.Windows.Forms.CheckBox();
            this._ttsCheckBox = new System.Windows.Forms.CheckBox();
            this._persistToasts = new System.Windows.Forms.CheckBox();
            this._enableLegacyToast = new System.Windows.Forms.CheckBox();
            this._disableToasts = new System.Windows.Forms.CheckBox();
            this._appTabControl = new System.Windows.Forms.TabControl();
            this._mainTabPage = new System.Windows.Forms.TabPage();
            this._mainTableLayout = new System.Windows.Forms.TableLayoutPanel();
            this._button1 = new System.Windows.Forms.Button();
            this._richTextBox1 = new System.Windows.Forms.RichTextBox();
            this._appTitle = new System.Windows.Forms.Label();
            this._copyrightLink = new System.Windows.Forms.LinkLabel();
            this._settingsPage = new System.Windows.Forms.TabPage();
            this._settingsPanel = new System.Windows.Forms.Panel();
            this._pushbulletSettings = new System.Windows.Forms.GroupBox();
            this._pushbulletDeviceIdTextBox = new System.Windows.Forms.TextBox();
            this._pushbulletDeviceIdlabel = new System.Windows.Forms.Label();
            this._pushbulletTokenTextBox = new System.Windows.Forms.TextBox();
            this._pushbulletTokenLabel = new System.Windows.Forms.Label();
            this._pushbulletCheckbox = new System.Windows.Forms.CheckBox();
            this._settingsTableLayout = new System.Windows.Forms.TableLayoutPanel();
            this._telegramSettings = new System.Windows.Forms.GroupBox();
            this._chatIdTextBox = new System.Windows.Forms.TextBox();
            this._ChatIdLabel = new System.Windows.Forms.Label();
            this._telegramTokenTextBox = new System.Windows.Forms.TextBox();
            this._tokenLabel = new System.Windows.Forms.Label();
            this._telegramCheckBox = new System.Windows.Forms.CheckBox();
            this._generalSettings = new System.Windows.Forms.GroupBox();
            this._testSettings = new System.Windows.Forms.GroupBox();
            this._toastSettings = new System.Windows.Forms.GroupBox();
            this._ttsSettings = new System.Windows.Forms.GroupBox();
            this._appTabControl.SuspendLayout();
            this._mainTabPage.SuspendLayout();
            this._mainTableLayout.SuspendLayout();
            this._settingsPage.SuspendLayout();
            this._settingsPanel.SuspendLayout();
            this._pushbulletSettings.SuspendLayout();
            this._settingsTableLayout.SuspendLayout();
            this._telegramSettings.SuspendLayout();
            this._generalSettings.SuspendLayout();
            this._testSettings.SuspendLayout();
            this._toastSettings.SuspendLayout();
            this._ttsSettings.SuspendLayout();
            this.SuspendLayout();
            // 
            // _label1
            // 
            this._label1.AutoSize = true;
            this._label1.Location = new System.Drawing.Point(3, 23);
            this._label1.Name = "_label1";
            this._label1.Size = new System.Drawing.Size(55, 13);
            this._label1.TabIndex = 0;
            this._label1.Text = "Language";
            // 
            // _languageValue
            // 
            this._languageValue.Location = new System.Drawing.Point(0, 0);
            this._languageValue.Name = "_languageValue";
            this._languageValue.Size = new System.Drawing.Size(100, 20);
            this._languageValue.TabIndex = 0;
            this._languageValue.TabStop = false;
            this._languageValue.Visible = false;
            // 
            // _languageComboBox
            // 
            this._languageComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this._languageComboBox.FormattingEnabled = true;
            this._languageComboBox.Location = new System.Drawing.Point(80, 23);
            this._languageComboBox.Name = "_languageComboBox";
            this._languageComboBox.Size = new System.Drawing.Size(130, 21);
            this._languageComboBox.TabIndex = 0;
            // 
            // _enableTestEnvironment
            // 
            this._enableTestEnvironment.AutoSize = true;
            this._enableTestEnvironment.Location = new System.Drawing.Point(6, 20);
            this._enableTestEnvironment.Name = "_enableTestEnvironment";
            this._enableTestEnvironment.Size = new System.Drawing.Size(145, 17);
            this._enableTestEnvironment.TabIndex = 5;
            this._enableTestEnvironment.Text = "Enable Test Environment";
            this._enableTestEnvironment.UseVisualStyleBackColor = true;
            // 
            // _ttsCheckBox
            // 
            this._ttsCheckBox.AutoSize = true;
            this._ttsCheckBox.Location = new System.Drawing.Point(6, 22);
            this._ttsCheckBox.Name = "_ttsCheckBox";
            this._ttsCheckBox.Size = new System.Drawing.Size(139, 17);
            this._ttsCheckBox.TabIndex = 4;
            this._ttsCheckBox.Text = "Enable Text To Speech";
            this._ttsCheckBox.UseVisualStyleBackColor = true;
            this._ttsCheckBox.CheckedChanged += new System.EventHandler(this._ttsCheckBox_CheckedChanged);
            // 
            // _persistToasts
            // 
            this._persistToasts.AutoSize = true;
            this._persistToasts.Location = new System.Drawing.Point(6, 45);
            this._persistToasts.Name = "_persistToasts";
            this._persistToasts.Size = new System.Drawing.Size(137, 17);
            this._persistToasts.TabIndex = 2;
            this._persistToasts.Text = "Make Toasts Persistent";
            this._persistToasts.UseVisualStyleBackColor = true;
            // 
            // _enableLegacyToast
            // 
            this._enableLegacyToast.AutoSize = true;
            this._enableLegacyToast.Location = new System.Drawing.Point(6, 68);
            this._enableLegacyToast.Name = "_enableLegacyToast";
            this._enableLegacyToast.Size = new System.Drawing.Size(132, 17);
            this._enableLegacyToast.TabIndex = 3;
            this._enableLegacyToast.Text = "Enable Legacy Toasts";
            this._enableLegacyToast.UseVisualStyleBackColor = true;
            // 
            // _disableToasts
            // 
            this._disableToasts.AutoSize = true;
            this._disableToasts.Location = new System.Drawing.Point(6, 22);
            this._disableToasts.Name = "_disableToasts";
            this._disableToasts.Size = new System.Drawing.Size(96, 17);
            this._disableToasts.TabIndex = 1;
            this._disableToasts.Text = "Disable Toasts";
            this._disableToasts.UseVisualStyleBackColor = true;
            this._disableToasts.CheckedChanged += new System.EventHandler(this._disableToasts_CheckedChanged);
            // 
            // _appTabControl
            // 
            this._appTabControl.Controls.Add(this._mainTabPage);
            this._appTabControl.Controls.Add(this._settingsPage);
            this._appTabControl.Dock = System.Windows.Forms.DockStyle.Fill;
            this._appTabControl.Location = new System.Drawing.Point(0, 0);
            this._appTabControl.Name = "_appTabControl";
            this._appTabControl.SelectedIndex = 0;
            this._appTabControl.Size = new System.Drawing.Size(948, 606);
            this._appTabControl.TabIndex = 0;
            this._appTabControl.TabStop = false;
            // 
            // _mainTabPage
            // 
            this._mainTabPage.Controls.Add(this._mainTableLayout);
            this._mainTabPage.Dock = System.Windows.Forms.DockStyle.Fill;
            this._mainTabPage.Location = new System.Drawing.Point(4, 22);
            this._mainTabPage.Name = "_mainTabPage";
            this._mainTabPage.Padding = new System.Windows.Forms.Padding(3);
            this._mainTabPage.Size = new System.Drawing.Size(940, 580);
            this._mainTabPage.TabIndex = 0;
            this._mainTabPage.Text = "Main";
            this._mainTabPage.ToolTipText = "Shows main info and logs";
            this._mainTabPage.UseVisualStyleBackColor = true;
            // 
            // _mainTableLayout
            // 
            this._mainTableLayout.ColumnCount = 3;
            this._mainTableLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this._mainTableLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this._mainTableLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this._mainTableLayout.Controls.Add(this._button1, 2, 0);
            this._mainTableLayout.Controls.Add(this._richTextBox1, 0, 1);
            this._mainTableLayout.Controls.Add(this._appTitle, 0, 0);
            this._mainTableLayout.Controls.Add(this._copyrightLink, 1, 0);
            this._mainTableLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this._mainTableLayout.Location = new System.Drawing.Point(3, 3);
            this._mainTableLayout.Name = "_mainTableLayout";
            this._mainTableLayout.RowCount = 2;
            this._mainTableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this._mainTableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this._mainTableLayout.Size = new System.Drawing.Size(934, 574);
            this._mainTableLayout.TabIndex = 0;
            // 
            // _button1
            // 
            this._button1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this._button1.AutoSize = true;
            this._button1.Location = new System.Drawing.Point(831, 3);
            this._button1.MinimumSize = new System.Drawing.Size(100, 25);
            this._button1.Name = "_button1";
            this._button1.Size = new System.Drawing.Size(100, 25);
            this._button1.TabIndex = 0;
            this._button1.Text = "Clear Logs";
            this._button1.UseVisualStyleBackColor = true;
            // 
            // _richTextBox1
            // 
            this._richTextBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this._mainTableLayout.SetColumnSpan(this._richTextBox1, 3);
            this._richTextBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this._richTextBox1.Location = new System.Drawing.Point(3, 34);
            this._richTextBox1.Name = "_richTextBox1";
            this._richTextBox1.ReadOnly = true;
            this._richTextBox1.Size = new System.Drawing.Size(928, 537);
            this._richTextBox1.TabIndex = 1;
            this._richTextBox1.Text = "";
            // 
            // _appTitle
            // 
            this._appTitle.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this._appTitle.AutoSize = true;
            this._appTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this._appTitle.Location = new System.Drawing.Point(3, 0);
            this._appTitle.Name = "_appTitle";
            this._appTitle.Size = new System.Drawing.Size(97, 31);
            this._appTitle.TabIndex = 2;
            this._appTitle.Text = "DFAssist ~";
            this._appTitle.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // _copyrightLink
            // 
            this._copyrightLink.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this._copyrightLink.AutoSize = true;
            this._copyrightLink.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this._copyrightLink.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline;
            this._copyrightLink.Location = new System.Drawing.Point(106, 0);
            this._copyrightLink.Name = "_copyrightLink";
            this._copyrightLink.Size = new System.Drawing.Size(107, 31);
            this._copyrightLink.TabIndex = 2;
            this._copyrightLink.TabStop = true;
            this._copyrightLink.Text = "© easly1989";
            this._copyrightLink.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // _settingsPage
            // 
            this._settingsPage.Controls.Add(this._settingsPanel);
            this._settingsPage.Dock = System.Windows.Forms.DockStyle.Fill;
            this._settingsPage.Location = new System.Drawing.Point(4, 22);
            this._settingsPage.Name = "_settingsPage";
            this._settingsPage.Padding = new System.Windows.Forms.Padding(3);
            this._settingsPage.Size = new System.Drawing.Size(940, 580);
            this._settingsPage.TabIndex = 1;
            this._settingsPage.Text = "Settings";
            this._settingsPage.ToolTipText = "Change Settings for DFAssist";
            this._settingsPage.UseVisualStyleBackColor = true;
            // 
            // _settingsPanel
            // 
            this._settingsPanel.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this._settingsPanel.AutoScroll = true;
            this._settingsPanel.Controls.Add(this._pushbulletSettings);
            this._settingsPanel.Controls.Add(this._settingsTableLayout);
            this._settingsPanel.Location = new System.Drawing.Point(0, 0);
            this._settingsPanel.Name = "_settingsPanel";
            this._settingsPanel.Size = new System.Drawing.Size(940, 580);
            this._settingsPanel.TabIndex = 0;
            // 
            // _pushbulletSettings
            // 
            this._pushbulletSettings.Controls.Add(this._pushbulletDeviceIdTextBox);
            this._pushbulletSettings.Controls.Add(this._pushbulletDeviceIdlabel);
            this._pushbulletSettings.Controls.Add(this._pushbulletTokenTextBox);
            this._pushbulletSettings.Controls.Add(this._pushbulletTokenLabel);
            this._pushbulletSettings.Controls.Add(this._pushbulletCheckbox);
            this._pushbulletSettings.Location = new System.Drawing.Point(4, 431);
            this._pushbulletSettings.Name = "_pushbulletSettings";
            this._pushbulletSettings.Size = new System.Drawing.Size(933, 100);
            this._pushbulletSettings.TabIndex = 1;
            this._pushbulletSettings.TabStop = false;
            this._pushbulletSettings.Text = "Pushbullet Settings";
            // 
            // _pushbulletDeviceIdTextBox
            // 
            this._pushbulletDeviceIdTextBox.Location = new System.Drawing.Point(110, 72);
            this._pushbulletDeviceIdTextBox.Name = "_pushbulletDeviceIdTextBox";
            this._pushbulletDeviceIdTextBox.Size = new System.Drawing.Size(390, 20);
            this._pushbulletDeviceIdTextBox.TabIndex = 9;
            // 
            // _pushbulletDeviceIdlabel
            // 
            this._pushbulletDeviceIdlabel.AutoSize = true;
            this._pushbulletDeviceIdlabel.Location = new System.Drawing.Point(6, 72);
            this._pushbulletDeviceIdlabel.Name = "_pushbulletDeviceIdlabel";
            this._pushbulletDeviceIdlabel.Size = new System.Drawing.Size(53, 13);
            this._pushbulletDeviceIdlabel.TabIndex = 8;
            this._pushbulletDeviceIdlabel.Text = "Device Id";
            // 
            // _pushbulletTokenTextBox
            // 
            this._pushbulletTokenTextBox.Location = new System.Drawing.Point(110, 43);
            this._pushbulletTokenTextBox.Name = "_pushbulletTokenTextBox";
            this._pushbulletTokenTextBox.Size = new System.Drawing.Size(390, 20);
            this._pushbulletTokenTextBox.TabIndex = 7;
            // 
            // _pushbulletTokenLabel
            // 
            this._pushbulletTokenLabel.AutoSize = true;
            this._pushbulletTokenLabel.Location = new System.Drawing.Point(6, 43);
            this._pushbulletTokenLabel.Name = "_pushbulletTokenLabel";
            this._pushbulletTokenLabel.Size = new System.Drawing.Size(76, 13);
            this._pushbulletTokenLabel.TabIndex = 6;
            this._pushbulletTokenLabel.Text = "Access Token";
            // 
            // _pushbulletCheckbox
            // 
            this._pushbulletCheckbox.AutoSize = true;
            this._pushbulletCheckbox.Location = new System.Drawing.Point(5, 19);
            this._pushbulletCheckbox.Name = "_pushbulletCheckbox";
            this._pushbulletCheckbox.Size = new System.Drawing.Size(172, 17);
            this._pushbulletCheckbox.TabIndex = 5;
            this._pushbulletCheckbox.Text = "Enable Pushbullet Notifications";
            this._pushbulletCheckbox.UseVisualStyleBackColor = true;
            this._pushbulletCheckbox.CheckedChanged += new System.EventHandler(this._pushbulletCheckbox_CheckedChanged);
            // 
            // _settingsTableLayout
            // 
            this._settingsTableLayout.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this._settingsTableLayout.AutoSize = true;
            this._settingsTableLayout.ColumnCount = 1;
            this._settingsTableLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this._settingsTableLayout.Controls.Add(this._telegramSettings, 0, 4);
            this._settingsTableLayout.Controls.Add(this._generalSettings, 0, 0);
            this._settingsTableLayout.Controls.Add(this._toastSettings, 0, 1);
            this._settingsTableLayout.Controls.Add(this._ttsSettings, 0, 2);
            this._settingsTableLayout.Location = new System.Drawing.Point(0, 3);
            this._settingsTableLayout.Name = "_settingsTableLayout";
            this._settingsTableLayout.RowCount = 5;
            this._settingsTableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this._settingsTableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this._settingsTableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this._settingsTableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this._settingsTableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this._settingsTableLayout.Size = new System.Drawing.Size(940, 424);
            this._settingsTableLayout.TabIndex = 0;
            // 
            // _telegramSettings
            // 
            this._telegramSettings.Controls.Add(this._chatIdTextBox);
            this._telegramSettings.Controls.Add(this._ChatIdLabel);
            this._telegramSettings.Controls.Add(this._telegramTokenTextBox);
            this._telegramSettings.Controls.Add(this._tokenLabel);
            this._telegramSettings.Controls.Add(this._telegramCheckBox);
            this._telegramSettings.Location = new System.Drawing.Point(3, 321);
            this._telegramSettings.Name = "_telegramSettings";
            this._telegramSettings.Size = new System.Drawing.Size(934, 100);
            this._telegramSettings.TabIndex = 1;
            this._telegramSettings.TabStop = false;
            this._telegramSettings.Text = "Telegram Settings";
            // 
            // _chatIdTextBox
            // 
            this._chatIdTextBox.Location = new System.Drawing.Point(111, 73);
            this._chatIdTextBox.Name = "_chatIdTextBox";
            this._chatIdTextBox.Size = new System.Drawing.Size(390, 20);
            this._chatIdTextBox.TabIndex = 4;
            // 
            // _ChatIdLabel
            // 
            this._ChatIdLabel.AutoSize = true;
            this._ChatIdLabel.Location = new System.Drawing.Point(7, 73);
            this._ChatIdLabel.Name = "_ChatIdLabel";
            this._ChatIdLabel.Size = new System.Drawing.Size(41, 13);
            this._ChatIdLabel.TabIndex = 3;
            this._ChatIdLabel.Text = "Chat Id";
            // 
            // _telegramTokenTextBox
            // 
            this._telegramTokenTextBox.Location = new System.Drawing.Point(111, 44);
            this._telegramTokenTextBox.Name = "_telegramTokenTextBox";
            this._telegramTokenTextBox.Size = new System.Drawing.Size(390, 20);
            this._telegramTokenTextBox.TabIndex = 2;
            // 
            // _tokenLabel
            // 
            this._tokenLabel.AutoSize = true;
            this._tokenLabel.Location = new System.Drawing.Point(7, 44);
            this._tokenLabel.Name = "_tokenLabel";
            this._tokenLabel.Size = new System.Drawing.Size(38, 13);
            this._tokenLabel.TabIndex = 1;
            this._tokenLabel.Text = "Token";
            // 
            // _telegramCheckBox
            // 
            this._telegramCheckBox.AutoSize = true;
            this._telegramCheckBox.Location = new System.Drawing.Point(6, 20);
            this._telegramCheckBox.Name = "_telegramCheckBox";
            this._telegramCheckBox.Size = new System.Drawing.Size(167, 17);
            this._telegramCheckBox.TabIndex = 0;
            this._telegramCheckBox.Text = "Enable Telegram Notifications";
            this._telegramCheckBox.UseVisualStyleBackColor = true;
            this._telegramCheckBox.CheckedChanged += new System.EventHandler(this._telegramCheckBox_CheckedChanged);
            // 
            // _generalSettings
            // 
            this._generalSettings.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this._generalSettings.Controls.Add(this._testSettings);
            this._generalSettings.Controls.Add(this._label1);
            this._generalSettings.Controls.Add(this._languageComboBox);
            this._generalSettings.Location = new System.Drawing.Point(3, 3);
            this._generalSettings.Name = "_generalSettings";
            this._generalSettings.Size = new System.Drawing.Size(934, 100);
            this._generalSettings.TabIndex = 0;
            this._generalSettings.TabStop = false;
            this._generalSettings.Text = "General Settings";
            // 
            // _testSettings
            // 
            this._testSettings.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this._testSettings.Controls.Add(this._enableTestEnvironment);
            this._testSettings.Location = new System.Drawing.Point(481, 0);
            this._testSettings.Name = "_testSettings";
            this._testSettings.Size = new System.Drawing.Size(457, 100);
            this._testSettings.TabIndex = 3;
            this._testSettings.TabStop = false;
            this._testSettings.Text = "Test Settings";
            // 
            // _toastSettings
            // 
            this._toastSettings.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this._toastSettings.Controls.Add(this._disableToasts);
            this._toastSettings.Controls.Add(this._enableLegacyToast);
            this._toastSettings.Controls.Add(this._persistToasts);
            this._toastSettings.Location = new System.Drawing.Point(3, 109);
            this._toastSettings.Name = "_toastSettings";
            this._toastSettings.Size = new System.Drawing.Size(934, 100);
            this._toastSettings.TabIndex = 1;
            this._toastSettings.TabStop = false;
            this._toastSettings.Text = "Toasts Settings";
            // 
            // _ttsSettings
            // 
            this._ttsSettings.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this._ttsSettings.Controls.Add(this._ttsCheckBox);
            this._ttsSettings.Location = new System.Drawing.Point(3, 215);
            this._ttsSettings.Name = "_ttsSettings";
            this._ttsSettings.Size = new System.Drawing.Size(934, 100);
            this._ttsSettings.TabIndex = 2;
            this._ttsSettings.TabStop = false;
            this._ttsSettings.Text = "Text To Speech Settings";
            // 
            // MainControl
            // 
            this.Controls.Add(this._appTabControl);
            this.Name = "MainControl";
            this.Size = new System.Drawing.Size(948, 606);
            this._appTabControl.ResumeLayout(false);
            this._mainTabPage.ResumeLayout(false);
            this._mainTableLayout.ResumeLayout(false);
            this._mainTableLayout.PerformLayout();
            this._settingsPage.ResumeLayout(false);
            this._settingsPanel.ResumeLayout(false);
            this._settingsPanel.PerformLayout();
            this._pushbulletSettings.ResumeLayout(false);
            this._pushbulletSettings.PerformLayout();
            this._settingsTableLayout.ResumeLayout(false);
            this._telegramSettings.ResumeLayout(false);
            this._telegramSettings.PerformLayout();
            this._generalSettings.ResumeLayout(false);
            this._generalSettings.PerformLayout();
            this._testSettings.ResumeLayout(false);
            this._testSettings.PerformLayout();
            this._toastSettings.ResumeLayout(false);
            this._toastSettings.PerformLayout();
            this._ttsSettings.ResumeLayout(false);
            this._ttsSettings.PerformLayout();
            this.ResumeLayout(false);

        }

        private void DisableToastsOnCheckedChanged(object sender, EventArgs e)
        {
            Logger.Debug($"[DisableToasts] Desired Value: {_disableToasts.Checked}!");
            _enableLegacyToast.Enabled = !_disableToasts.Checked;
            _persistToasts.Enabled = _enableLegacyToast.Enabled && !_enableLegacyToast.Checked;
        }

        private void EnableLegacyToastsOnCheckedChanged(object sender, EventArgs e)
        {
            Logger.Debug($"[LegacyToasts] Desired Value: {_enableLegacyToast.Checked}!");
            _persistToasts.Enabled = !_enableLegacyToast.Checked;
            ToastWindowNotification(Localization.GetText("ui-toast-notification-test-title"), Localization.GetText("ui-toast-notification-test-message"));
        }

        private void PersistToastsOnCheckedChanged(object sender, EventArgs e)
        {
            try
            {
                Logger.Debug($"[PersistentToasts] Desired Value: {_persistToasts.Checked}!");

                var keyName = $@"Software\Microsoft\Windows\CurrentVersion\Notifications\Settings\{AppId}";
                using(var key = Registry.CurrentUser.OpenSubKey(keyName, true))
                {
                    if(_persistToasts.Checked)
                    {
                        if(key == null)
                        {
                            Logger.Debug("[PersistentToasts] Key not found in the registry, Adding a new one!");
                            Registry.SetValue($@"HKEY_CURRENT_USER\{keyName}", "ShowInActionCenter", 1, RegistryValueKind.DWord);
                        }
                        else
                        {
                            Logger.Debug("[PersistentToasts] Key found in the registry, setting value to 1!");
                            key.SetValue("ShowInActionCenter", 1, RegistryValueKind.DWord);
                        }
                    }
                    else
                    {
                        if(key == null)
                        {
                            Logger.Debug("[PersistentToasts] Key not found in the registry, nothing to do!");
                            return;
                        }

                        Logger.Debug("[PersistentToasts] Key found in the registry, Removing value!");
                        key.DeleteValue("ShowInActionCenter");
                    }

                    MessageBox.Show(Localization.GetText("ui-persistent-toast-warning-message"), Localization.GetText("ui-persistent-toast-warning-title"), MessageBoxButtons.OK);
                }
            }
            catch(Exception ex)
            {
                Logger.Exception(ex, "Unable to remove/add the registry key to make Toasts persistent!");
            }
        }

        #endregion

        #region IActPluginV1 Implementations

        public void InitPlugin(TabPage pluginScreenSpace, Label pluginStatusText)
        {
            _labelStatus = pluginStatusText;
            _labelTab = pluginScreenSpace;

            if(_mainFormIsLoaded)
                OnInit();
            else
                ActGlobals.oFormActMain.Shown += ActMainFormOnShown;
        }

        private void FormActMain_UpdateCheckClicked()
        {
            const int pluginId = 71;
            try
            {
                var localDate = ActGlobals.oFormActMain.PluginGetSelfDateUtc(this);
                var remoteDate = ActGlobals.oFormActMain.PluginGetRemoteDateUtc(pluginId);
                if(localDate.AddHours(2) >= remoteDate)
                    return;

                var result = MessageBox.Show(Localization.GetText("ui-update-available-message"),
                    Localization.GetText("ui-update-available-title"), MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);
                if(result != DialogResult.Yes)
                    return;

                var updatedFile = ActGlobals.oFormActMain.PluginDownload(pluginId);
                var pluginData = ActGlobals.oFormActMain.PluginGetSelfData(this);
                if(pluginData.pluginFile.Directory != null)
                    ActGlobals.oFormActMain.UnZip(updatedFile.FullName, pluginData.pluginFile.Directory.FullName);

                ThreadInvokes.CheckboxSetChecked(ActGlobals.oFormActMain, pluginData.cbEnabled, false);
                Application.DoEvents();
                ThreadInvokes.CheckboxSetChecked(ActGlobals.oFormActMain, pluginData.cbEnabled, true);
            }
            catch(Exception ex)
            {
                ActGlobals.oFormActMain.WriteExceptionLog(ex, "Plugin Update Check");
            }
        }

        private ActPluginData _ffxivPlugin;
        private void OnInit()
        {
            if(_pluginInitializing)
                return;

            _pluginInitializing = true;

            if(_ffxivPlugin != null)
                _ffxivPlugin.cbEnabled.CheckedChanged -= FFXIVParsingPlugin_IsEnabledChanged;

            // Before anything else, if the FFXIV Parsing Plugin is not already initialized
            // than this plugin cannot start
            var plugins = ActGlobals.oFormActMain.ActPlugins;
            _ffxivPlugin = plugins.FirstOrDefault(x => x.lblPluginTitle.Text == "FFXIV_ACT_Plugin.dll");
            if(_ffxivPlugin == null)
            {
                _pluginInitializing = false;
                ActGlobals.oFormActMain.PluginGetSelfData(this).cbEnabled.Checked = false;
                _labelStatus.Text = "FFXIV_ACT_Plugin must be installed BEFORE DFAssist!";
                return;
            }
            else
            {
                _ffxivPlugin.cbEnabled.CheckedChanged += FFXIVParsingPlugin_IsEnabledChanged;
                if(!_ffxivPlugin.cbEnabled.Checked)
                {
                    _pluginInitializing = false;
                    ActGlobals.oFormActMain.PluginGetSelfData(this).cbEnabled.Checked = false;
                    _labelStatus.Text = "FFXIV_ACT_Plugin must be enabled";
                    return;
                }
            }

            ActGlobals.oFormActMain.Shown -= ActMainFormOnShown;

            var pluginData = ActGlobals.oFormActMain.PluginGetSelfData(this);
            var enviroment = Path.GetDirectoryName(pluginData.pluginFile.ToString());
            AssemblyResolver.Initialize(enviroment);

            Logger.SetTextBox(_richTextBox1);
            Logger.Debug("----------------------------------------------------------------");
            Logger.Debug("Plugin Init");
            Logger.Debug($"Plugin Version: {Assembly.GetExecutingAssembly().GetName().Version}");

            var defaultLanguage = new Language { Name = "English", Code = "en-us" };
            LoadData(defaultLanguage);

            // The shortcut must be created to work with windows 8/10 Toasts
            Logger.Debug(ShortCutCreator.TryCreateShortcut(AppId, AppId)
                ? "Shortcut for ACT found"
                : "Unable to Create the Shorctut for ACT");

            _isPluginEnabled = true;

            Logger.Debug("Plugin Enabled");

            _languageComboBox.DataSource = new[]
            {
                defaultLanguage,
                new Language {Name = "한국어", Code = "ko-kr"},
                new Language {Name = "日本語", Code = "ja-jp"},
                new Language {Name = "Français", Code = "fr-fr"}
            };
            _languageComboBox.DisplayMember = "Name";
            _languageComboBox.ValueMember = "Code";
            

            _labelStatus.Text = "Starting...";

            _labelStatus.Text = Localization.GetText("l-plugin-started");
            _labelTab.Text = Localization.GetText("app-name");

            Logger.Debug("Plugin Started!");

            _labelTab.Controls.Add(this);
            _xmlSettingsSerializer = new SettingsSerializer(this);

            LoadSettings();
            LoadData();
            UpdateProcesses();

            _languageComboBox.SelectedValueChanged += LanguageComboBox_SelectedValueChanged;

            if(_timer == null)
            {
                _timer = new Timer { Interval = 30000 };
                _timer.Tick += Timer_Tick;
            }

            _timer.Enabled = true;            
            _pluginInitializing = false;

            ActGlobals.oFormActMain.UpdateCheckClicked += FormActMain_UpdateCheckClicked;
            if(ActGlobals.oFormActMain.GetAutomaticUpdatesAllowed())
                new Thread(FormActMain_UpdateCheckClicked).Start();

            Logger.Debug("----------------------------------------------------------------");
        }

        private void FFXIVParsingPlugin_IsEnabledChanged(object sender, EventArgs e)
        {
            if(!_ffxivPlugin.cbEnabled.Checked)
            {
                ActGlobals.oFormActMain.PluginGetSelfData(this).cbEnabled.Checked = false;
                DeInitPlugin();
            }
        }

        public void DeInitPlugin()
        {
            if(!_isPluginEnabled)
                return;

            _isPluginEnabled = false;

            if(_ffxivPlugin != null)
                _ffxivPlugin.cbEnabled.CheckedChanged -= FFXIVParsingPlugin_IsEnabledChanged;

            SaveSettings();

            _labelTab = null;

            if(_labelStatus != null)
            {
                _labelStatus.Text = Localization.GetText("l-plugin-stopped");
                _labelStatus = null;
            }

            foreach(var entry in _networks)
                entry.Value.Network.StopCapture();

            _timer.Enabled = false;

            Logger.SetTextBox(null);
        }

        #endregion

        #region Getters

        private static string GetInstanceName(int code)
        {
            return Data.GetInstance(code).Name;
        }

        private static string GetRouletteName(int code)
        {
            return Data.GetRoulette(code).Name;
        }

        #endregion

        #region Update Methods

        private void UpdateProcesses()
        {
            var process = Process.GetProcessesByName("ffxiv_dx11").FirstOrDefault();
            if(process == null)
                return;
            try
            {
                if(!_networks.ContainsKey(process.Id))
                {
                    var pn = new ProcessNet(process, new Network());
                    FFXIVPacketHandler.OnEventReceived += Network_onReceiveEvent;
                    _networks.TryAdd(process.Id, pn);
                    Logger.Success("l-process-set-success", process.Id);
                }
            }
            catch(Exception e)
            {
                Logger.Exception(e, "l-process-set-failed");
            }

            var toDelete = new List<int>();
            foreach(var entry in _networks)
            {
                if(entry.Value.Process.HasExited)
                {
                    entry.Value.Network.StopCapture();
                    toDelete.Add(entry.Key);
                }
                else
                {
                    if(entry.Value.Network.IsRunning)
                        entry.Value.Network.UpdateGameConnections(entry.Value.Process);
                    else
                    {
                        if(!entry.Value.Network.StartCapture(entry.Value.Process))
                            toDelete.Add(entry.Key);
                    }
                }
            }

            foreach(var t in toDelete)
            {
                try
                {
                    _networks.TryRemove(t, out _);
                    FFXIVPacketHandler.OnEventReceived -= Network_onReceiveEvent;
                }
                catch(Exception e)
                {
                    Logger.Exception(e, "l-process-remove-failed");
                }
            }
        }

        private void UpdateTranslations()
        {
            Logger.Debug("Updating Localization for UI...");

            _label1.Text = Localization.GetText("ui-language-display-text");
            _button1.Text = Localization.GetText("ui-log-clear-display-text");
            _enableTestEnvironment.Text = Localization.GetText("ui-enable-test-environment");
            _ttsCheckBox.Text = Localization.GetText("ui-enable-tts");
            _persistToasts.Text = Localization.GetText("ui-persist-toasts");
            _enableLegacyToast.Text = Localization.GetText("ui-enable-legacy-toasts");
            _disableToasts.Text = Localization.GetText("ui-disable-toasts");
            _appTitle.Text = $"{Localization.GetText("app-name")} v{Assembly.GetExecutingAssembly().GetName().Version} | ";
            _generalSettings.Text = Localization.GetText("ui-general-settings-group");
            _toastSettings.Text = Localization.GetText("ui-toast-settings-group");
            _ttsSettings.Text = Localization.GetText("ui-tts-settings-group");
            _testSettings.Text = Localization.GetText("ui-test-settings-group");
            _telegramSettings.Text = Localization.GetText("ui-telegram-settings-group");
            _telegramCheckBox.Text = Localization.GetText("ui-telegram-display-text");
            _tokenLabel.Text = Localization.GetText("ui-telegram-token-display-text");
            _ChatIdLabel.Text = Localization.GetText("ui-telegram-chatid-display-text");
            _telegramSettings.Text = Localization.GetText("ui-telegram-settings-group");
            _pushbulletSettings.Text = Localization.GetText("ui-pushbullet-settings-group");
            _pushbulletCheckbox.Text = Localization.GetText("ui-pushbullet-display-text");
            _pushbulletDeviceIdlabel.Text = Localization.GetText("ui-pushbullet-deviceid-display-text");
            _pushbulletTokenLabel.Text = Localization.GetText("ui-pushbullet-token-display-text");

            Logger.Debug("Localization for UI Updated!");
        }

        #endregion

        #region Post Method

        private static void SendToAct(string text)
        {
            ActGlobals.oFormActMain.ParseRawLogLine(false, DateTime.Now, "00|" + DateTime.Now.ToString("O") + "|0048|F|" + text);
        }

        private void PostToToastWindowsNotificationIfNeeded(string server, EventType eventType, int[] args)
        {
            if(eventType != EventType.MATCH_ALERT)
                return;

            var head = _networks.Count <= 1 ? "" : "[" + server + "] ";
            var title = head + (args[0] != 0 ? GetRouletteName(args[0]) : Localization.GetText("app-name"));
            var testing = _enableTestEnvironment.Checked ? "[Code: " + args[1] + "] " : string.Empty;

            switch (eventType)
            {
                case EventType.MATCH_ALERT:                    
                    SendPushbulletNotification(title, ">> " + GetInstanceName(args[1]));
                    SendTelegramNotification(title, ">> " + GetInstanceName(args[1]));
                    ToastWindowNotification(title, ">> " + testing + GetInstanceName(args[1]));
                    TtsNotification(GetInstanceName(args[1]));
                    break;
                case EventType.MATCH_ORDER_PROGRESS:
                    var order = args[1];

                    if (order == 1)
                    {
                        var message = ">> " + string.Format(Localization.GetText("l-queue-order"), order);
                        SendTelegramNotification(title, message, true);                        
                    }
                    break;
            }
        }

        private void SendPushbulletNotification(string title, string messageText)
        {
            Logger.Debug("Request received to send Pushbullet notification");
            if (!_pushbulletCheckbox.Checked)
            {
                Logger.Debug("Pushbullet notifications are disabled");
                return;
            }

            if (string.IsNullOrWhiteSpace(_pushbulletTokenTextBox.Text))
            {
                Logger.Debug("Pushbullet Token is missing.");
                return;
            }

            try
            {
                PushbulletClient client = new PushbulletClient(_pushbulletTokenTextBox.Text);
                PushbulletSharp.Models.Requests.PushNoteRequest request = new PushbulletSharp.Models.Requests.PushNoteRequest() { Body = messageText, Title = title };
                if (!string.IsNullOrWhiteSpace(_pushbulletDeviceIdTextBox.Text)) request.DeviceIden = _pushbulletDeviceIdTextBox.Text;

                var response = client.PushNote(request);

                Logger.Debug("Message pushed to Pushbullet with Id " + response.ReceiverIden);
            }
            catch (Exception ex)
            {
                Logger.Exception(ex, "Unable to push Pushbullet notification.");
                Logger.Debug(ex.ToString());
            }
        }

        private void SendTelegramNotification(string title, string messageText, bool disableNotification = false)
        {
            Logger.Debug("Request received to send Telegram notification");
            if(!_telegramCheckBox.Checked)
            {
                Logger.Debug("Telegram notifications are disabled");
                return;
            }

            if(string.IsNullOrWhiteSpace(_telegramTokenTextBox.Text))
            {
                Logger.Debug("Token is missing.");
                return;
            }

            if(string.IsNullOrWhiteSpace(_chatIdTextBox.Text))
            {
                Logger.Debug("Chat Id is missing");
                return;
            }

            try
            {
                var botClient = new TelegramBotClient(_telegramTokenTextBox.Text);
                Telegram.Bot.Types.ChatId chatId;
                int chatIdInt = 0;
                long chatIdentifier = 0;

                if (int.TryParse(_chatIdTextBox.Text, out chatIdInt))
                    chatId = new Telegram.Bot.Types.ChatId(chatIdInt);
                else if (long.TryParse(_chatIdTextBox.Text, out chatIdentifier))
                    chatId = new Telegram.Bot.Types.ChatId(chatIdentifier);
                else
                    chatId = new Telegram.Bot.Types.ChatId(_chatIdTextBox.Text);

                var message = botClient.SendTextMessageAsync(chatId, title + " " + messageText, disableNotification: disableNotification).Result;
                Logger.Debug($"Telegram notification sent with message Id {message.MessageId}");
            }
            catch(Exception ex)
            {
                Logger.Exception(ex, "Unable to send Telegram notification.");
                Logger.Debug(ex.ToString());
            }
        }

        private void ToastWindowNotification(string title, string message)
        {
            Logger.Debug("Request Showing Taost received...");
            if(_disableToasts.Checked)
            {
                Logger.Debug("... Toasts are disabled!");
                return;
            }

            if(_enableLegacyToast.Checked)
            {
                Logger.Debug("... Legacy Toasts Enabled...");
                try
                {
                    Logger.Debug("... Closing any open Legacy Toast...");
                    _lastToast?.Close();
                    LegacyToastDispose();
                    Application.ThreadException += LegacyToastOnGuiUnhandedException;
                    AppDomain.CurrentDomain.UnhandledException += LegacyToastOnUnhandledException;
                    var toast = new LegacyToast(title, message, _networks) { Text = title };
                    Logger.Debug("... Creating new Legacy Toast...");
                    _lastToast = toast;
                    _lastToast.Closing += LastToastOnClosing;
                    _lastToast.Show();
                    Logger.Debug("... Legacy Toast Showing...");
                    NativeMethods.ShowWindow(_lastToast.Handle, 9);
                    NativeMethods.SetForegroundWindow(_lastToast.Handle);
                    _lastToast.Activate();
                }
                catch(Exception ex)
                {
                    Logger.Debug("Error handling/creating Legacy Toast!");
                    LegacyToastHandleUnhandledException(ex);
                    _lastToast?.Close();
                    LegacyToastDispose();
                }
            }
            else
            {
                Logger.Debug("... Legacy Toasts Disabled...");
                try
                {
                    var toastXml = ToastNotificationManager.GetTemplateContent(ToastTemplateType.ToastImageAndText03);

                    var stringElements = toastXml.GetElementsByTagName("text");
                    if(stringElements.Length < 2)
                    {
                        Logger.Error("l-toast-notification-error");
                        return;
                    }

                    stringElements[0].AppendChild(toastXml.CreateTextNode(title));
                    stringElements[1].AppendChild(toastXml.CreateTextNode(message));

                    var toast = new ToastNotification(toastXml);
                    Logger.Debug("... Creating new Toast...");
                    ToastNotificationManager.CreateToastNotifier(AppId).Show(toast);
                    Logger.Debug("... Toast Showing...");
                }
                catch(Exception e)
                {
                    Logger.Exception(e, "l-toast-notification-error");
                }
            }
        }

        private void LegacyToastDispose()
        {
            Application.ThreadException -= LegacyToastOnGuiUnhandedException;
            AppDomain.CurrentDomain.UnhandledException -= LegacyToastOnUnhandledException;
            if(_lastToast == null || _lastToast.IsDisposed)
                return;
            _lastToast.Closing -= LastToastOnClosing;
            _lastToast.Dispose();
        }

        #endregion

        #region Events

        private void LastToastOnClosing(object sender, CancelEventArgs cancelEventArgs)
        {
            LegacyToastDispose();
        }

        private static void LegacyToastOnUnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            LegacyToastHandleUnhandledException(e.ExceptionObject as Exception);
        }

        private static void LegacyToastOnGuiUnhandedException(object sender, ThreadExceptionEventArgs e)
        {
            LegacyToastHandleUnhandledException(e.Exception);
        }

        private static void LegacyToastHandleUnhandledException(Exception e)
        {
            if(e == null)
                return;
            Logger.Exception(e, "l-toast-notification-error");
        }

        private void EnableTtsOnCheckedChanged(object sender, EventArgs eventArgs)
        {
            
        }

        private void TtsNotification(string message, string title = "ui-tts-dutyfound")
        {
            if(!_ttsCheckBox.Checked)
                return;

            var dutyFound = Localization.GetText(title);
            _synth.Speak(dutyFound); // duty found
            _synth.Speak(message);
        }

        private void Network_onReceiveEvent(int pid, EventType eventType, int[] args)
        {
            var server = _networks[pid].Process.MainModule.FileName.Contains("KOREA") ? "KOREA" : "GLOBAL";
            var text = pid + "|" + server + "|" + eventType + "|";
            var pos = 0;

            switch(eventType)
            {
                case EventType.MATCH_ALERT:
                    text += GetRouletteName(args[0]) + "|";
                    pos++;
                    text += GetInstanceName(args[1]) + "|";
                    pos++;
                    break;
            }

            for(var i = pos; i < args.Length; i++)
                text += args[i] + "|";

            SendToAct(text);

            PostToToastWindowsNotificationIfNeeded(server, eventType, args);
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            if(!_isPluginEnabled)
                return;

            UpdateProcesses();
        }

        private void ClearLogsButton_Click(object sender, EventArgs e)
        {
            _richTextBox1.Clear();
        }

        private void ActMainFormOnShown(object sender, EventArgs e)
        {
            _mainFormIsLoaded = true;
            OnInit();
        }

        private void LanguageComboBox_SelectedValueChanged(object sender, EventArgs e)
        {
            LoadData();
            UpdateTranslations();
        }

        #endregion

        #region Settings

        private void LoadSettings()
        {
            Logger.Debug("Settings Loading...");
            // All the settings to deserialize
            _xmlSettingsSerializer.AddControlSetting(_disableToasts.Name, _disableToasts);
            _xmlSettingsSerializer.AddControlSetting(_languageValue.Name, _languageValue);
            _xmlSettingsSerializer.AddControlSetting(_ttsCheckBox.Name, _ttsCheckBox);
            _xmlSettingsSerializer.AddControlSetting(_persistToasts.Name, _persistToasts);
            _xmlSettingsSerializer.AddControlSetting(_enableTestEnvironment.Name, _enableTestEnvironment);
            _xmlSettingsSerializer.AddControlSetting(_enableLegacyToast.Name, _enableLegacyToast);
            _xmlSettingsSerializer.AddControlSetting(_telegramCheckBox.Name, _telegramCheckBox);
            _xmlSettingsSerializer.AddControlSetting(_telegramTokenTextBox.Name, _telegramTokenTextBox);
            _xmlSettingsSerializer.AddControlSetting(_chatIdTextBox.Name, _chatIdTextBox);
            _xmlSettingsSerializer.AddControlSetting(_pushbulletCheckbox.Name, _pushbulletCheckbox);
            _xmlSettingsSerializer.AddControlSetting(_pushbulletDeviceIdTextBox.Name, _pushbulletDeviceIdTextBox);
            _xmlSettingsSerializer.AddControlSetting(_pushbulletTokenTextBox.Name, _pushbulletTokenTextBox);

            if(File.Exists(_settingsFile))
                using(var fileStream = new FileStream(_settingsFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using(var xmlTextReader = new XmlTextReader(fileStream))
                {
                    try
                    {
                        while(xmlTextReader.Read())
                        {
                            if(xmlTextReader.NodeType != XmlNodeType.Element)
                                continue;

                            if(xmlTextReader.LocalName == "SettingsSerializer")
                                _xmlSettingsSerializer.ImportFromXml(xmlTextReader);
                        }
                    }
                    catch(Exception ex)
                    {
                        _labelStatus.Text = Localization.GetText("l-settings-load-error", ex.Message);
                    }

                    xmlTextReader.Close();
                }

            foreach(var language in _languageComboBox.Items.OfType<Language>())
            {
                if(language.Name.Equals(_languageValue.Text))
                {
                    _languageComboBox.SelectedItem = language;
                }
            }

            Logger.Debug($"Language: {_languageValue.Text}");
            Logger.Debug($"Disable Toasts: {_disableToasts.Checked}");
            Logger.Debug($"Make Toasts Persistent: {_persistToasts.Checked}");
            Logger.Debug($"Enable Legacy Toasts: {_enableLegacyToast.Checked}");
            Logger.Debug($"Enable Text To Speech: {_ttsCheckBox.Checked}");
            Logger.Debug($"Enable Test Environment: {_enableTestEnvironment.Checked}");
            Logger.Debug("Settings Loaded!");
        }

        private void SaveSettings()
        {
            try
            {
                Logger.Debug("Saving Settings...");
                using(var fileStream = new FileStream(_settingsFile, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
                using(var xmlTextWriter = new XmlTextWriter(fileStream, Encoding.UTF8) { Formatting = Formatting.Indented, Indentation = 1, IndentChar = '\t' })
                {
                    xmlTextWriter.WriteStartDocument(true);
                    xmlTextWriter.WriteStartElement("Config"); // <Config>
                    xmlTextWriter.WriteStartElement("SettingsSerializer"); // <Config><SettingsSerializer>
                    _xmlSettingsSerializer.ExportToXml(xmlTextWriter); // Fill the SettingsSerializer XML
                    xmlTextWriter.WriteEndElement(); // </SettingsSerializer>
                    xmlTextWriter.WriteEndElement(); // </Config>
                    xmlTextWriter.WriteEndDocument(); // Tie up loose ends (shouldn't be any)
                    xmlTextWriter.Flush(); // Flush the file buffer to disk
                    xmlTextWriter.Close();

                    Logger.Debug("Settings Saved!");
                }
            }
            catch(Exception ex)
            {
                Logger.Exception(ex, "l-settings-save-error");
            }
        }

        #endregion

        private void _ttsCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            Logger.Debug($"[TTS] Desired Value: {_ttsCheckBox.Checked}!");

            if(_ttsCheckBox.Checked)
                TtsNotification(Localization.GetText("ui-tts-notification-test-message"), Localization.GetText("ui-tts-notification-test-title"));
        }

        private void _telegramCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            Logger.Debug($"[Telegram] Desired Value: {_telegramCheckBox.Checked}!");

            if (_telegramCheckBox.Checked)
                SendTelegramNotification(Localization.GetText("ui-toast-notification-test-title"), Localization.GetText("ui-telegram-notification-test-message"));
        }

        private void _pushbulletCheckbox_CheckedChanged(object sender, EventArgs e)
        {
            Logger.Debug($"[Pushbullet] Desired Value: {_pushbulletCheckbox.Checked}!");

            if (_pushbulletCheckbox.Checked)
                SendPushbulletNotification(Localization.GetText("ui-toast-notification-test-title"), Localization.GetText("ui-pushbullet-notification-test-message"));
        }

        private void _disableToasts_CheckedChanged(object sender, EventArgs e)
        {
            Logger.Debug($"[Toasts] Desired Value: {!_ttsCheckBox.Checked}!");

            if (!_disableToasts.Checked)
                ToastWindowNotification(Localization.GetText("ui-toast-notification-test-title"), Localization.GetText("ui-toast-notification-test-message"));
        }
    }
}