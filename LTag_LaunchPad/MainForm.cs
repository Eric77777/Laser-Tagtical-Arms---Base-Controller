using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.IO.Ports; //Serial port code
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace LTag_LaunchPad
{
    public partial class MainForm : Form
    {

        public List<String> PlayerNames = new List<String>();
        public List<int> PlayerIDs = new List<int>();
        public List<decimal> WinPerc = new List<decimal>();
        public List<decimal> AccuracyPerc = new List<decimal>();
        public List<decimal> KillDeathRatio = new List<decimal>();
        public String serialCom = ""; //String for incoming serial communications
        public String serialComOut = "";  //String for outgoing serial communications
        public int serialComInt; //Serial In Msg: Msg Type
        public int serialComInt1; //Serial In Msg: Msg Type
        public int serialComInt2; //Serial In Msg: Variable Index
        public int serialComInt3; //Serial In Msg: Variable Value
        public bool setupComplete = false; //Indicator for Form Load class being complete
        public int stepSelected = 1; //Enables user to progress through game management steps
        public int tabSelected = 0; //Enables restricting user from selecting tabs, instead use buttons
        public int dataGridViewRowCounter = 0; //Enables counting amongst functions (global variable)
        private delegate void BatteryCheckInDelegate(int UnitRow);


        public MainForm()
        {

            InitializeComponent(); //Automatically generated method that contains all form components and properties

        }



        private void MainForm_Load(object sender, EventArgs e)
        {


            #region GATHER SERIAL PORT INFO AND POPULATE COMBOBOXES
            string[] ports = SerialPort.GetPortNames();
            ControlBaseSerialportTitle_comboBox.Items.Clear();
            foreach (string comport in ports)  //Add all port options to combobox, then set first as the default
            {
                ControlBaseSerialportTitle_comboBox.Items.Add(comport);
            }
            if (ControlBaseSerialportTitle_comboBox.Items.Count > 0)
            {
                ControlBaseSerialportTitle_comboBox.SelectedIndex = 0;  //Default Serial Port.  Efficient.
            }
            ControlBaseRateTitle_comboBox.SelectedItem = "9600";  //Default Serial Port Baud Rate.
            setupComplete = true; //Enable events
            #endregion



            GameSetup_AddUnitID_comboBox.SelectedItem = "1";  //Default UNIT ID.  Needed.



            #region INITIAL SETUP OF SERIAL PORT

            serialPort1.BaudRate = Convert.ToInt32(Convert.ToString(ControlBaseRateTitle_comboBox.Text));
            serialPort1.Parity = Parity.None;
            serialPort1.DataBits = 8;
            serialPort1.StopBits = StopBits.One;
            serialPort1.RtsEnable = true;
            serialPort1.Handshake = Handshake.None;
            serialPort1.NewLine = "\n";

            serialPort1.DataReceived += new SerialDataReceivedEventHandler(serialPort1_DataReceived);

            if (ControlBaseSerialportTitle_comboBox.Text != null && ControlBaseSerialportTitle_comboBox.Text != "")
            {
                serialPort1.PortName = Convert.ToString(ControlBaseSerialportTitle_comboBox.Text);

                RadioCheck();
            }

            #endregion



            #region IMPORT PLAYER DATA

            //Open Excel and gather info about data limits
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Lasertag\Historical Player Data\HistoricalData.xlsx");
            Excel.Worksheet xlWorksheet = xlWorkbook.Sheets["CompiledStats"];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;


            for (int i = 2; i <= rowCount; i++) //Start at row 2 of worksheet in order to skip header row
            {
                PlayerNames.Add(xlRange.Cells[i, 1].Value2);
                PlayerIDs.Add(Convert.ToInt16(xlRange.Cells[i, 2].Value2));
                WinPerc.Add(Convert.ToDecimal(xlRange.Cells[i, 3].Value2));
                AccuracyPerc.Add(Convert.ToDecimal(xlRange.Cells[i, 4].Value2));
                KillDeathRatio.Add(Convert.ToDecimal(xlRange.Cells[i, 5].Value2));
            }



            //CLEANUP AND CLOSE ALL EXCEL ELEMENTS - BEGIN
            //CLEANUP
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //RELEASE COM OBJECTS TO FULLY KILL EXCELL PROCESS FROM RUNNING IN THE BACKGROUND
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //CLOSE AND RELEASE
            xlWorkbook.Close(0);
            Marshal.ReleaseComObject(xlWorkbook);

            //QUIT AND RELEASE
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            //CLEANUP AND CLOSE ALL EXCEL ELEMENTS - END
            #endregion



            #region BIND DATA SOURCES (INTERNALLY DEFINED VARIABLES AND EXTERNAL DATA SOURCES) TO MAINFORM
            this.GameSetup_GameType_comboBox.DataSource = VariableDefinition.GameType; //Assign GameType variable as datasource for GameType combobox
            this.GameSetup_GameTime_comboBox.DataSource = VariableDefinition.GameTime2; //Assign GameTime1 variable as datasource for GameTime combobox
            this.GameSetup_PregameTime_comboBox.DataSource = VariableDefinition.PregameTime; //Assign PregameTime variable as datasource for PregameTime combobox
            //this.GameSetup_datagridView_PlayerName.DataSource = PlayerList; //Bind list to Game Setup datagridview (object variable is assigned as datasource in the datagridview properties)
            this.GameSetup_AddPlayer_comboBox.DataSource = PlayerNames; //Bind list AddPlayer combobox (PlayerListCurrent keeps running condition of unselected players)
            #endregion



        }



        #region ALL STEPS - STEP FORWARD AND BACK THROUGH GAME SETUP STEPS
        private void StepForward_button_Click(object sender, EventArgs e)
        {
            if (stepSelected == 1)
            {
                RadioCheck();
                RosterPlayerCount();
                if (StepStatus1_textBox.Text == "OK" & StepStatus2_textBox.Text == "OK" & StepStatus3_textBox.Text == "OK")
                {
                    InitiateStep2();
                    stepSelected = 2;
                }
            }
            else if (stepSelected == 2)
            {
                if (StepStatus1_textBox.Text == "OK" & StepStatus2_textBox.Text == "OK" & StepStatus3_textBox.Text == "OK")
                {
                    InitiateStep3();
                    stepSelected = 3;
                }
            }
            else if (stepSelected == 3)
            {
                if (StepStatus1_textBox.Text == "OK" & StepStatus2_textBox.Text == "OK" & StepStatus3_textBox.Text == "OK")
                {
                    InitiateStep4();
                    stepSelected = 4;
                }
            }
            else if (stepSelected == 4)
            {
                if (StepStatus1_textBox.Text == "OK" & StepStatus2_textBox.Text == "OK" & StepStatus3_textBox.Text == "OK")
                {
                    InitiateStep5();
                    stepSelected = 5;
                }
            }
        }



        private void StepBackward_button_Click(object sender, EventArgs e)
        {
            if (stepSelected == 2)
            {
                InitiateStep1();
                stepSelected = 1;
            }
            else if (stepSelected == 3)
            {
                InitiateStep2();
                stepSelected = 2;
            }
            else if (stepSelected == 4)
            {
                InitiateStep3();
                stepSelected = 3;
            }
            else if (stepSelected == 5)
            {
                InitiateStep4();
                stepSelected = 4;
            }
        }



        private void InitiateStep1() //Game Setup (1st Tab)
        {
            GameSetup_panel1.Enabled = true;
            GameSetup_panel2.Enabled = true;
            ControlBaseRadio_panel.Enabled = true;

            GameSetup_AddPlayer_button.Enabled = true; //UNLOCK FOR ADDING
            GameSetup_AddNewPlayer_button.Enabled = true; //UNLOCK FOR ADDING
            GameSetup_datagridView.AllowUserToDeleteRows = true; //ENABLE DELETING WHEN IN STEP 1

            StepBackward_button.Visible = false;
            StepBackward_button.Enabled = false;
            StepForward_button.Visible = true;
            StepForward_button.Enabled = true;
            StepTitle_button.Text = "STEP 1" + Environment.NewLine + "Roster";
            StepTitle_textBox.Text = "STEP 1 - Roster Setup & Control Base Radio Check";
            StepDirections_textBox.Text = "ROSTER - Add all relevant Players and Settings (Need > 1 Player)." + Environment.NewLine + "RADIO - Select valid SERIAL PORT and TX/RX RATE.";

            StepStatus1Title_textBox.Visible = true;
            StepStatus1Title_textBox.Text = "ROSTER:";
            StepStatus1_textBox.Visible = true;
            RosterPlayerCount();

            StepStatus2Title_textBox.Visible = true;
            StepStatus2Title_textBox.Text = "RADIO:";
            StepStatus2_textBox.Visible = true;
            RadioCheck();

            StepStatus3Title_textBox.Visible = false;
            StepStatus3Title_textBox.Text = "NONE";
            StepStatus3_textBox.Visible = false;
            StepStatus3_textBox.Text = "OK";
            StepStatus3_textBox.BackColor = System.Drawing.Color.Lime;

            tabSelected = 0;  //Display first tab
            this.Invoke(new EventHandler(Game_tabControl_SelectedIndexChanged));

            GameSetup_datagridView_UnitID.Visible = true;
            GameSetup_datagridView_PlayerName.Visible = true;
            GameSetup_datagridView_PlayerID.Visible = true;
            GameSetup_datagridView_PlayerWinPerc.Visible = false;
            GameSetup_datagridView_PlayerAccuracyPerc.Visible = false;
            GameSetup_datagridView_KillDeathRatio.Visible = false;
            GameSetup_datagridView_GunBattery.Visible = false;
            GameSetup_datagridView_VestBattery.Visible = false;
            GameSetup_datagridView_Team.Visible = true;
            GameSetup_datagridView_Team.ReadOnly = false;
            GameSetup_datagridView_Health.Visible = true;
            GameSetup_datagridView_Shield.Visible = true;
            GameSetup_datagridView_GunROF.Visible = true;
            GameSetup_datagridView_GunDamage.Visible = true;
            GameSetup_datagridView_Ammo.Visible = true;
            GameSetup_datagridView_SaveData.Visible = true;
            GameSetup_datagridView_BattCheck.Visible = false;
            GameSetup_datagridView_UnitBoot.Visible = false;
            GameSetup_datagridView_Shots.Visible = false;
            GameSetup_datagridView_Hits.Visible = false;
            GameSetup_datagridView_DamageDone.Visible = false;
            GameSetup_datagridView_Eliminations.Visible = false;
            GameSetup_datagridView_Wounds.Visible = false;
            GameSetup_datagridView_DamageTaken.Visible = false;
            GameSetup_datagridView_Survived.Visible = false;
            GameSetup_datagridView_Revives.Visible = false;
            GameSetup_datagridView_Points.Visible = false;
            GameSetup_datagridView_Placement.Visible = false;
        }



        private void InitiateStep2() //Game Boot (1st Tab)
        {
            GameSetup_panel1.Enabled = false;
            GameSetup_panel2.Enabled = false;
            ControlBaseRadio_panel.Enabled = false;

            GameSetup_AddPlayer_button.Enabled = false; //LOCK FROM ADDING
            GameSetup_AddNewPlayer_button.Enabled = false; //UNLOCK FROM ADDING
            GameSetup_datagridView.AllowUserToDeleteRows = false; //DISABLE DELETING WHEN NOT IN STEP 1

            StepBackward_button.Visible = true;
            StepBackward_button.Enabled = true;
            StepForward_button.Visible = true;
            StepForward_button.Enabled = true;

            StepTitle_button.Text = "STEP 2" + Environment.NewLine + "Batteries";
            StepTitle_textBox.Text = "STEP 2 - Unit battery checks";
            StepDirections_textBox.Text = "BATTERIES - Verify unit batteries > 50%";

            StepStatus1Title_textBox.Visible = true;
            StepStatus1Title_textBox.Text = "BATTERIES:";
            StepStatus1_textBox.Visible = true;
            UnitBatteryCheckAll();

            StepStatus2Title_textBox.Visible = false;
            StepStatus2Title_textBox.Text = "NONE";
            StepStatus2_textBox.Visible = false;
            StepStatus3_textBox.Text = "OK";
            StepStatus3_textBox.BackColor = System.Drawing.Color.Lime;

            StepStatus3Title_textBox.Visible = false;
            StepStatus3Title_textBox.Text = "NONE";
            StepStatus3_textBox.Visible = false;
            StepStatus3_textBox.Text = "OK";
            StepStatus3_textBox.BackColor = System.Drawing.Color.Lime;


            tabSelected = 0;  //Display first tab
            this.Invoke(new EventHandler(Game_tabControl_SelectedIndexChanged));

            GameSetup_datagridView_UnitID.Visible = true;
            GameSetup_datagridView_PlayerName.Visible = true;
            GameSetup_datagridView_PlayerID.Visible = true;
            GameSetup_datagridView_PlayerWinPerc.Visible = false;
            GameSetup_datagridView_PlayerAccuracyPerc.Visible = false;
            GameSetup_datagridView_KillDeathRatio.Visible = false;
            GameSetup_datagridView_GunBattery.Visible = true;
            GameSetup_datagridView_VestBattery.Visible = true;
            GameSetup_datagridView_Team.Visible = true;
            GameSetup_datagridView_Team.ReadOnly = true;
            GameSetup_datagridView_Health.Visible = false;
            GameSetup_datagridView_Shield.Visible = false;
            GameSetup_datagridView_GunROF.Visible = false;
            GameSetup_datagridView_GunDamage.Visible = false;
            GameSetup_datagridView_Ammo.Visible = false;
            GameSetup_datagridView_SaveData.Visible = false;
            GameSetup_datagridView_BattCheck.Visible = true;
            GameSetup_datagridView_UnitBoot.Visible = false;
            GameSetup_datagridView_Shots.Visible = false;
            GameSetup_datagridView_Hits.Visible = false;
            GameSetup_datagridView_DamageDone.Visible = false;
            GameSetup_datagridView_Eliminations.Visible = false;
            GameSetup_datagridView_Wounds.Visible = false;
            GameSetup_datagridView_DamageTaken.Visible = false;
            GameSetup_datagridView_Survived.Visible = false;
            GameSetup_datagridView_Revives.Visible = false;
            GameSetup_datagridView_Points.Visible = false;
            GameSetup_datagridView_Placement.Visible = false;
        }



        private void InitiateStep3() //Game Boot (1st Tab)
        {
            GameSetup_panel1.Enabled = false;
            GameSetup_panel2.Enabled = false;
            ControlBaseRadio_panel.Enabled = false;

            GameSetup_AddPlayer_button.Enabled = false; //LOCK FROM ADDING
            GameSetup_AddNewPlayer_button.Enabled = false; //UNLOCK FROM ADDING
            GameSetup_datagridView.AllowUserToDeleteRows = false; //DISABLE DELETING WHEN NOT IN STEP 1

            StepBackward_button.Visible = true;
            StepBackward_button.Enabled = true;
            StepForward_button.Visible = true;
            StepForward_button.Enabled = true;

            StepTitle_button.Text = "STEP 3" + Environment.NewLine + "Game Boot";
            StepTitle_textBox.Text = "STEP 3 - Boot Units";
            StepDirections_textBox.Text = "BOOT - Verify game boot to all units";

            StepStatus1Title_textBox.Visible = true;
            StepStatus1Title_textBox.Text = "BOOT:";
            StepStatus1_textBox.Visible = true;
            //BootGameCheckAll();

            StepStatus2Title_textBox.Visible = false;
            StepStatus2Title_textBox.Text = "NONE";
            StepStatus2_textBox.Visible = false;
            StepStatus2_textBox.Text = "OK";
            StepStatus2_textBox.BackColor = System.Drawing.Color.Lime;

            StepStatus3Title_textBox.Visible = false;
            StepStatus3Title_textBox.Text = "NONE";
            StepStatus3_textBox.Visible = false;
            StepStatus3_textBox.Text = "OK";
            StepStatus3_textBox.BackColor = System.Drawing.Color.Lime;


            tabSelected = 0;  //Display first tab
            this.Invoke(new EventHandler(Game_tabControl_SelectedIndexChanged));

            GameSetup_datagridView_UnitID.Visible = true;
            GameSetup_datagridView_PlayerName.Visible = true;
            GameSetup_datagridView_PlayerID.Visible = true;
            GameSetup_datagridView_PlayerWinPerc.Visible = false;
            GameSetup_datagridView_PlayerAccuracyPerc.Visible = false;
            GameSetup_datagridView_KillDeathRatio.Visible = false;
            GameSetup_datagridView_GunBattery.Visible = true;
            GameSetup_datagridView_VestBattery.Visible = true;
            GameSetup_datagridView_Team.Visible = true;
            GameSetup_datagridView_Team.ReadOnly = true;
            GameSetup_datagridView_Health.Visible = true;
            GameSetup_datagridView_Shield.Visible = true;
            GameSetup_datagridView_GunROF.Visible = true;
            GameSetup_datagridView_GunDamage.Visible = true;
            GameSetup_datagridView_Ammo.Visible = true;
            GameSetup_datagridView_SaveData.Visible = false;
            GameSetup_datagridView_BattCheck.Visible = false;
            GameSetup_datagridView_UnitBoot.Visible = true;
            GameSetup_datagridView_Shots.Visible = false;
            GameSetup_datagridView_Hits.Visible = false;
            GameSetup_datagridView_DamageDone.Visible = false;
            GameSetup_datagridView_Eliminations.Visible = false;
            GameSetup_datagridView_Wounds.Visible = false;
            GameSetup_datagridView_DamageTaken.Visible = false;
            GameSetup_datagridView_Survived.Visible = false;
            GameSetup_datagridView_Revives.Visible = false;
            GameSetup_datagridView_Points.Visible = false;
            GameSetup_datagridView_Placement.Visible = false;
        }



        private void InitiateStep4() //Game Start and Running (2nd Tab)
        {
            GameSetup_panel1.Enabled = false;
            GameSetup_panel2.Enabled = false;
            ControlBaseRadio_panel.Enabled = false;

            GameSetup_AddPlayer_button.Enabled = false; //LOCK FROM ADDING
            GameSetup_AddNewPlayer_button.Enabled = false; //UNLOCK FROM ADDING
            GameSetup_datagridView.AllowUserToDeleteRows = false; //DISABLE DELETING WHEN NOT IN STEP 1

            StepBackward_button.Visible = true;
            StepBackward_button.Enabled = true;
            StepForward_button.Visible = true;
            StepForward_button.Enabled = true;

            StepTitle_button.Text = "STEP 4" + Environment.NewLine + "Gametime";
            StepTitle_textBox.Text = "STEP 4 - Game Start and End";
            StepDirections_textBox.Text = "Start - Game started" + Environment.NewLine + "End - Game ended";

            StepStatus1Title_textBox.Visible = true;
            StepStatus1Title_textBox.Text = "START:";
            StepStatus1_textBox.Visible = true;
            //GameStarted();

            StepStatus2Title_textBox.Visible = true;
            StepStatus2Title_textBox.Text = "END:";
            StepStatus2_textBox.Visible = true;
            //GameEnded();

            StepStatus3Title_textBox.Visible = false;
            StepStatus3Title_textBox.Text = "NONE";
            StepStatus3_textBox.Visible = false;
            StepStatus3_textBox.Text = "OK";
            StepStatus3_textBox.BackColor = System.Drawing.Color.Lime;

            tabSelected = 1;  //Progress to second tab
            this.Invoke(new EventHandler(Game_tabControl_SelectedIndexChanged));
        }



        private void InitiateStep5() //Game End and Gather Data(1st Tab)
        {
            GameSetup_panel1.Enabled = false;
            GameSetup_panel2.Enabled = false;
            ControlBaseRadio_panel.Enabled = false;

            GameSetup_AddPlayer_button.Enabled = false; //LOCK FROM ADDING
            GameSetup_AddNewPlayer_button.Enabled = false; //UNLOCK FROM ADDING
            GameSetup_datagridView.AllowUserToDeleteRows = false; //DISABLE DELETING WHEN NOT IN STEP 1

            StepBackward_button.Visible = true;
            StepBackward_button.Enabled = true;
            StepForward_button.Visible = false;
            StepForward_button.Enabled = false;
            StepTitle_button.Text = "Step 4" + Environment.NewLine + "- Gather Data -";

            StepTitle_button.Text = "STEP 5" + Environment.NewLine + "Gather Data";
            StepTitle_textBox.Text = "STEP 5 - Gather game data";
            StepDirections_textBox.Text = "DATA: Verify all unit data gathered";

            StepStatus1Title_textBox.Visible = true;
            StepStatus1Title_textBox.Text = "DATA:";
            StepStatus1_textBox.Visible = true;
            //GameDataGathered();

            StepStatus2Title_textBox.Visible = false;
            StepStatus2Title_textBox.Text = "NONE";
            StepStatus2_textBox.Visible = false;
            StepStatus2_textBox.Text = "OK";
            StepStatus2_textBox.BackColor = System.Drawing.Color.Lime;

            StepStatus3Title_textBox.Visible = false;
            StepStatus3Title_textBox.Text = "NONE";
            StepStatus3_textBox.Visible = false;
            StepStatus3_textBox.Text = "OK";
            StepStatus3_textBox.BackColor = System.Drawing.Color.Lime;

            tabSelected = 0;  //Progress to first tab
            this.Invoke(new EventHandler(Game_tabControl_SelectedIndexChanged));

            GameSetup_datagridView_UnitID.Visible = true;
            GameSetup_datagridView_PlayerName.Visible = true;
            GameSetup_datagridView_PlayerID.Visible = true;
            GameSetup_datagridView_PlayerWinPerc.Visible = false;
            GameSetup_datagridView_PlayerAccuracyPerc.Visible = false;
            GameSetup_datagridView_KillDeathRatio.Visible = false;
            GameSetup_datagridView_GunBattery.Visible = false;
            GameSetup_datagridView_VestBattery.Visible = false;
            GameSetup_datagridView_Team.Visible = true;
            GameSetup_datagridView_Health.Visible = true;
            GameSetup_datagridView_Shield.Visible = true;
            GameSetup_datagridView_GunROF.Visible = true;
            GameSetup_datagridView_GunDamage.Visible = true;
            GameSetup_datagridView_Ammo.Visible = true;
            GameSetup_datagridView_SaveData.Visible = true;
            GameSetup_datagridView_BattCheck.Visible = false;
            GameSetup_datagridView_UnitBoot.Visible = false;
            GameSetup_datagridView_Shots.Visible = true;
            GameSetup_datagridView_Hits.Visible = true;
            GameSetup_datagridView_DamageDone.Visible = true;
            GameSetup_datagridView_Eliminations.Visible = true;
            GameSetup_datagridView_Wounds.Visible = true;
            GameSetup_datagridView_DamageTaken.Visible = true;
            GameSetup_datagridView_Survived.Visible = true;
            GameSetup_datagridView_Revives.Visible = true;
            GameSetup_datagridView_Points.Visible = true;
            GameSetup_datagridView_Placement.Visible = true;
        }
        #endregion



        #region STEP 1 - RADIO CHECK FUNCTIONS AND EVENTS
        private void RadioStatus_button_Click(object sender, EventArgs e)
        {
            if (ControlBaseSerialportTitle_comboBox.Text != null && ControlBaseSerialportTitle_comboBox.Text != "")
            {

                if (serialPort1.IsOpen == true) { serialPort1.Close(); }

                serialPort1.PortName = Convert.ToString(ControlBaseSerialportTitle_comboBox.Text);

                RadioCheck();
            }
            else
            {
                this.Invoke(new EventHandler(RadioCheck_NOK));
            }
        }



        private void SerialPort_comboBox_SelectedValueChanged(object sender, EventArgs e)
        {
            if (ControlBaseSerialportTitle_comboBox.Text != null && ControlBaseSerialportTitle_comboBox.Text != "" && setupComplete == true)
            {
                if (serialPort1.IsOpen == true) { serialPort1.Close(); }

                serialPort1.PortName = Convert.ToString(ControlBaseSerialportTitle_comboBox.Text);

                RadioCheck();
            }
            else if (Convert.ToString(ControlBaseSerialportTitle_comboBox.Text) == null)
            {
                this.Invoke(new EventHandler(RadioCheck_NOK));
            }

        }



        private void SerialRate_comboBox_SelectedValueChanged(object sender, EventArgs e)
        {
            if (ControlBaseSerialportTitle_comboBox.Text != null && ControlBaseSerialportTitle_comboBox.Text != "" && setupComplete == true)
            {
                if (serialPort1.IsOpen == true) { serialPort1.Close(); }

                serialPort1.BaudRate = Convert.ToInt32(Convert.ToString(ControlBaseRateTitle_comboBox.Text));

                RadioCheck();
            }
            else if (Convert.ToString(ControlBaseSerialportTitle_comboBox.Text) == null)
            {
                this.Invoke(new EventHandler(RadioCheck_NOK));
            }
        }



        private void RadioCheck()
        {
            setupComplete = false; //temp disable event

            if (ControlBaseSerialportTitle_comboBox.Text != "" && ControlBaseSerialportTitle_comboBox.Text != null)
            {
                string serialPortName = ControlBaseSerialportTitle_comboBox.Text;

                bool portNameExists = false;

                string[] ports = SerialPort.GetPortNames();
                ControlBaseSerialportTitle_comboBox.Items.Clear();

                foreach (string comport in ports)
                {
                    ControlBaseSerialportTitle_comboBox.Items.Add(comport);
                    if (comport == serialPortName) //Verify that selected port is still available/connected
                    {
                        portNameExists = true;
                    }
                }

                if (portNameExists == true)
                {
                    ControlBaseSerialportTitle_comboBox.Text = serialPortName;
                    try
                    {
                        if (serialPort1.IsOpen == false) { serialPort1.Open(); }

                        serialPort1.Write("1000000000\r");
                    }
                    catch (IOException)
                    {

                    }
                }
                else
                {
                    ControlBaseSerialportTitle_comboBox.Text = serialPortName;
                    this.Invoke(new EventHandler(RadioCheck_NOK));
                }
            }
            setupComplete = true;
        }



        private void RadioCheck_OK(object sender, EventArgs e)
        {
            if (stepSelected == 1)
            {
                StepStatus2_textBox.Text = "OK";
                StepStatus2_textBox.BackColor = System.Drawing.Color.Lime;
                serialCom = "";
            }
        }



        private void RadioCheck_NOK(object sender, EventArgs e)
        {
            if (stepSelected == 1)
            {
                StepStatus2_textBox.Text = "NOK";
                StepStatus2_textBox.BackColor = System.Drawing.Color.Red;
                serialCom = "";
            }
        }



        private void SerialPort_comboBox_DropDown(object sender, EventArgs e)
        {
            string[] ports = SerialPort.GetPortNames();
            ControlBaseSerialportTitle_comboBox.Items.Clear();
            foreach (string comport in ports)
            {
                ControlBaseSerialportTitle_comboBox.Items.Add(comport);
            }

            if (ControlBaseSerialportTitle_comboBox.Items.Count > 0)
            {
                ControlBaseSerialportTitle_comboBox.SelectedIndex = 0;
            }

        }
        #endregion



        #region STEP 1 - GAME TYPE SELECTED ON GameSetup TAB
        //COMBOBOX - GAME SELECTION AND DIRECTIONS - BEGIN---------------------------------------------------------------------------------------
        private void GamecomboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (GameSetup_GameType_comboBox.Text == "Free for All")
            {
                Setup_GameSpan_radioButton1.Enabled = true; //Survival game type enabled (not selected)
                Setup_GameSpan_radioButton2.Enabled = true; //Timed game type enabled (not selected)
                Setup_GameSpan_radioButton1.Checked = true;
                Setup_GameDirections_textBox.Text = "A simple game for all players. There are no teams and everyone is fighting to either be the last player standing (survival), or the player with the most points at the end of the game (timed).";
            }
            else if (GameSetup_GameType_comboBox.Text == "Teams/Squads")
            {
                Setup_GameSpan_radioButton1.Enabled = true; //Survival game type enabled (not selected)
                Setup_GameSpan_radioButton2.Enabled = true; //Timed game type enabled (not selected)
                Setup_GameSpan_radioButton1.Checked = true;
                Setup_GameDirections_textBox.Text = "Teams face off to eliminate (Elimination) or out-score the opposition (Timed).";
            }
            else if (GameSetup_GameType_comboBox.Text == "Survivor")
            {
                Setup_GameSpan_radioButton1.Checked = true; //Survival game enabled, only
                Setup_GameSpan_radioButton1.Enabled = false; //Survival game type disabled
                Setup_GameSpan_radioButton2.Enabled = false; //Timed game type disabled
                GameSetup_GameTime_comboBox.Text = "N/A";  //Time set to "N/A"
                GameSetup_GameTime_comboBox.Enabled = false; //Time disabled
                Setup_GameDirections_textBox.Text = "A game that begins with Teams/Squads and morphs into Free-for-All when more than half of the players have been eliminated.";
            }
        }
        //COMBOBOX - GAME SELECTION AND DIRECTIONS - END^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
        #endregion



        #region STEP 1 - GAME SPAN SELECTED ON GameSetup TAB
        //RADIOBUTTON - GAME TYPE - BEGIN-----------------------------------------------------------------------------------------
        private void GameTyperadioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (Setup_GameSpan_radioButton1.Checked)
            {
                this.GameSetup_GameTime_comboBox.DataSource = VariableDefinition.GameTime2; //List without times assigned
                GameSetup_GameTime_comboBox.Text = "N/A";
                GameSetup_GameTime_comboBox.Enabled = false;
            }
            else if (Setup_GameSpan_radioButton2.Checked)
            {
                this.GameSetup_GameTime_comboBox.DataSource = VariableDefinition.GameTime1; //List with times assigned
                GameSetup_GameTime_comboBox.Enabled = true;
                GameSetup_GameTime_comboBox.Text = "20";
            }
        }
        //RADIOBUTTON - GAME TYPE - END^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
        #endregion

        

        #region STEP 1 - ADD PRIOR PLAYER BUTTON CLICK
        private void GameSetup_AddPlayer_button_Click(object sender, EventArgs e)
        {
            bool NameExists = false;
            bool UnitIDExists = false;
            string SelectedPlayer = this.GameSetup_AddPlayer_comboBox.GetItemText(this.GameSetup_AddPlayer_comboBox.SelectedItem);
            string SelectedUnitID = this.GameSetup_AddUnitID_comboBox.GetItemText(this.GameSetup_AddUnitID_comboBox.SelectedItem);

            //CHECK TO SEE IF PLAYER IS ALREADY ON ROSTER
            if (GameSetup_datagridView.RowCount > 0 && GameSetup_AddPlayer_comboBox.SelectedValue != null)
            {
                for (int i = 0; i < GameSetup_datagridView.RowCount; i++)
                {
                    if (String.Equals(GameSetup_datagridView.Rows[i].Cells[1].Value.ToString(), SelectedPlayer))
                    {
                        NameExists = true;
                    }
                    if (String.Equals(GameSetup_datagridView.Rows[i].Cells[0].Value.ToString(), SelectedUnitID))
                    {
                        UnitIDExists = true;
                    }
                }
            }

            if (NameExists == true || UnitIDExists == true)  //NAME OR UNIT ID IS ALREADY ON ROSTER
            {
                if (NameExists == true)
                {
                    MessageBox.Show("PLAYER NAME ALREADY IN USE");
                }
                if (UnitIDExists == true)
                {
                    MessageBox.Show("UNIT ID ALREADY IN USE ON ROSTER");
                }
            }
            else //NAME IS NOT YET ON ROSTER, ADD IT
            {
                int NameIndex = PlayerNames.IndexOf(GameSetup_AddPlayer_comboBox.Text);

                this.GameSetup_datagridView.Rows.Insert(GameSetup_datagridView.RowCount, SelectedUnitID, PlayerNames[NameIndex], PlayerIDs[NameIndex],
                    WinPerc[NameIndex], AccuracyPerc[NameIndex], KillDeathRatio[NameIndex], VariableDefinition.defaultGunBattery,
                    VariableDefinition.defaultVestBattery, VariableDefinition.defaultTeam, VariableDefinition.defaultHealth,
                    VariableDefinition.defaultShield, VariableDefinition.defaultGunROF, VariableDefinition.defaultGunDamage,
                    VariableDefinition.defaultAmmo, VariableDefinition.defaultSave, VariableDefinition.BattCheck, VariableDefinition.BootUnit);

                RosterPlayerCount(); //CHECK AND CHANGE STATUS OF ROSTER
            }
        }
        #endregion



        #region STEP 1 - ADD NEW PLAYER BUTTON CLICK
        private void GameSetup_AddNewPlayer_button_Click(object sender, EventArgs e)
        {
            bool NameExists = false;
            bool UnitIDExists = false;
            string SelectedPlayer = GameSetup_AddNewPlayer_textBox.Text.ToString();
            string SelectedUnitID = this.GameSetup_AddUnitID_comboBox.GetItemText(this.GameSetup_AddUnitID_comboBox.SelectedItem);
            int TempPlayerID = 0;
            int MaxPlayerID = 0;

            //CHECK TO SEE IF PLAYER ALREADY EXISTS IN HISTORICAL DATA, AND CHECK FOR HIGHEST PLAYERID IN HISTORICAL DATA
            if (GameSetup_AddNewPlayer_textBox.Text != "" && GameSetup_AddPlayer_comboBox.SelectedValue != null)
            {
                for (int i = 0; i < PlayerNames.Count; i++)
                {
                    if (String.Equals(GameSetup_AddNewPlayer_textBox.Text.ToString(), PlayerNames[i],StringComparison.OrdinalIgnoreCase))
                    {
                        NameExists = true;
                    }
                    if (MaxPlayerID < PlayerIDs[i])
                    {
                        MaxPlayerID = PlayerIDs[i];
                    }
                }

                //NOW CHECK IF ANY NEW PLAYERS HAVE ALREADY BEEN ADDED TO ROSTER, THUS INCREASING THE HIGHEST ASSIGNED PLAYERID
                if (GameSetup_datagridView.RowCount > 0)
                {
                    for (int i = 0; i < GameSetup_datagridView.RowCount; i++)
                    {
                        if (String.Equals(GameSetup_datagridView.Rows[i].Cells[1].Value.ToString(), SelectedPlayer, StringComparison.OrdinalIgnoreCase))
                        {
                            NameExists = true;
                        }
                        if (String.Equals(GameSetup_datagridView.Rows[i].Cells[0].Value.ToString(), SelectedUnitID))
                        {
                            UnitIDExists = true;
                        }
                        TempPlayerID = Convert.ToInt16(GameSetup_datagridView.Rows[i].Cells[2].Value);
                        if (MaxPlayerID < TempPlayerID)
                        {
                            MaxPlayerID = TempPlayerID;
                        }

                    }
                }

                if (NameExists == true || UnitIDExists == true)  //NAME OR UNIT ID IS ALREADY ON ROSTER
                {
                    if (NameExists == true)
                    {
                        MessageBox.Show("PLAYER NAME ALREADY IN USE");
                    }
                    if (UnitIDExists == true)
                    {
                        MessageBox.Show("UNIT ID ALREADY IN USE ON ROSTER");
                    }
                }
                else  //NAME IS NOT YET ON ROSTER, ADD IT
                {
                    int NameIndex = PlayerNames.IndexOf(GameSetup_AddPlayer_comboBox.Text);

                    this.GameSetup_datagridView.Rows.Insert(GameSetup_datagridView.RowCount, SelectedUnitID, GameSetup_AddNewPlayer_textBox.Text, MaxPlayerID + 1,
                        "----", "----", "----",
                        VariableDefinition.defaultGunBattery, VariableDefinition.defaultVestBattery, VariableDefinition.defaultTeam,
                        VariableDefinition.defaultHealth, VariableDefinition.defaultShield, VariableDefinition.defaultGunROF,
                        VariableDefinition.defaultGunDamage, VariableDefinition.defaultAmmo, VariableDefinition.defaultSave,
                        VariableDefinition.BattCheck, VariableDefinition.BootUnit);

                    RosterPlayerCount(); //CHECK AND CHANGE STATUS OF ROSTER
                }
            }
        }
        #endregion



        #region STEP 1 EXIT CRITERION - CHECK PLAYER ROSTER COUNT
        private void RosterPlayerCount()
        {
            if (GameSetup_datagridView.RowCount > 1)
            {
                StepStatus1_textBox.Text = "OK";
                StepStatus1_textBox.BackColor = System.Drawing.Color.Lime;
            }
            else
            {
                StepStatus1_textBox.Text = "NOK";
                StepStatus1_textBox.BackColor = System.Drawing.Color.Red;
            }
        }
        #endregion



        #region STEP 2 & 3 - DATAGRIDVIEW CELL CLICK EVENT
        private void GameSetup_datagridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (GameSetup_datagridView.Columns[e.ColumnIndex].Name == "GameSetup_datagridView_BattCheck")
            {
                dataGridViewRowCounter = GameSetup_datagridView.Rows[e.RowIndex].Index;
                UnitBatteryCheck();
            }
            if (GameSetup_datagridView.Columns[e.ColumnIndex].Name == "GameSetup_datagridView_UnitBoot")
            {
                //THIS AREA SAVED FOR REBOOTING SPECIFIC UNITS (DATAGRIDVIEW ROW)

                // Here you call your method that deals with the row values
                // you can use e.RowIndex to find the row

                // I'm also just leaving item as type object but since you control the form
                // you can usually safely cast to a specific object here.
            }
        }
        #endregion



        #region STEP 2 EXIT CRITERION - CHECK UNIT BATTERIES, DATAGRIDVIEW CELL CLICK EVENT
        private void UnitBatteryCheckAll()
        {
            for (dataGridViewRowCounter = 0; dataGridViewRowCounter < GameSetup_datagridView.RowCount; dataGridViewRowCounter++)
            {
                UnitBatteryCheck();
            }
        }

        private void UnitBatteryCheck()
        {
            //dataGridViewRowCounter assigned either in UnitBatteryCheckAll or Battery Check button click (TBD) of DataGridView
            for (int TaggerVestIndicator = 0; TaggerVestIndicator <= 1; TaggerVestIndicator++) //Battery Check for both Tagger and Vest
            {
                try
                {
                    //GameSetup_datagridView.Rows[dataGridViewRowCounter].Cells[6 + TaggerVestIndicator].Value = "0"; //Set initial value
                    serialComOut = Convert.ToString(2000000000 + (TaggerVestIndicator * 100000000) + Convert.ToInt32(GameSetup_datagridView.Rows[dataGridViewRowCounter].Cells[0].Value.ToString()) * 1000000) + "\r";

                    if (serialPort1.IsOpen == false) { serialPort1.Open(); }

                    Debug.WriteLine(serialComOut);

                    serialPort1.Write(serialComOut);
                }
                catch (IOException)
                {

                }
            }
        }
        #endregion



        #region STEP 3 EXIT CRITERION - BOOT UNITS
        private void BootGameCheckAll()
        {

        }
        #endregion



        #region STEP 4 EXIT CRITERION - CHECK GAME ENDED
        private void GameStarted()
        {

        }
        private void GameEnded()
        {

        }
        #endregion



        #region STEP 5 EXIT CRITERION - GAME DATA GATHERED
        private void GameDataGathered()
        {

        }
        #endregion



        #region ALL STEPS - SHOW/HIDE PLAYER DATA COLUMNS
        private void ShowHidePlayerData_button_Click(object sender, EventArgs e)
        {
            if (GameSetup_datagridView_PlayerWinPerc.Visible == false)
            {
                GameSetup_datagridView_PlayerWinPerc.Visible = true;
                GameSetup_datagridView_PlayerAccuracyPerc.Visible = true;
                GameSetup_datagridView_KillDeathRatio.Visible = true;
            }
            else
            {
                GameSetup_datagridView_PlayerWinPerc.Visible = false;
                GameSetup_datagridView_PlayerAccuracyPerc.Visible = false;
                GameSetup_datagridView_KillDeathRatio.Visible = false;
            }
        }
        #endregion



        #region SERIAL PORT DATA RECEIVED
        public void serialPort1_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            serialCom = serialPort1.ReadLine();

            Debug.WriteLine(serialCom);

            Int32.TryParse(serialCom, out serialComInt);
            serialComInt1 = Convert.ToInt16(serialComInt / 1000000);
            serialComInt2 = Convert.ToInt16((serialComInt - serialComInt1 * 1000000)/1000);
            serialComInt3 = Convert.ToInt16(serialComInt - serialComInt1 * 1000000 - serialComInt2 * 1000);

            if (serialComInt1 == 1) //RADIO CHECK
            {
                this.Invoke(new EventHandler(RadioCheck_OK));
            }
            else if (serialComInt1 == 2) //BATTERY CHECKS
            {
                BatteryCheckIn();
            }
        }

        public void BatteryCheckIn ()
        {
            int UnitRow = 0;
            foreach (DataGridViewRow row in GameSetup_datagridView.Rows) //Find row with matching Unit ID
            {
                if (Convert.ToInt16(row.Cells[0].Value) == serialComInt2 - 100 * Convert.ToInt16(Math.Floor(Convert.ToDecimal(serialComInt2) / 100)))
                {
                    UnitRow = row.Index;
                    break;
                }
            }
            GameSetup_datagridView.Invoke(new BatteryCheckInDelegate(BatteryPercent), UnitRow); //Delegate
        }

        private void BatteryPercent(int UnitRow) //Delegate
        {
            if (serialComInt3 <= 100) //message from Unit (not error)
            {
                GameSetup_datagridView.Rows[UnitRow].Cells[6 + Convert.ToInt16(Math.Floor(Convert.ToDecimal(serialComInt2) / 100))].Value = Convert.ToDecimal(serialComInt3) / 100; //Divide by 100 for %
                if (serialComInt3 <= 50)
                {
                    GameSetup_datagridView.Rows[UnitRow].Cells[6 + Convert.ToInt16(Math.Floor(Convert.ToDecimal(serialComInt2) / 100))].Style.BackColor = Color.Red;
                }
                else
                {
                    GameSetup_datagridView.Rows[UnitRow].Cells[6 + Convert.ToInt16(Math.Floor(Convert.ToDecimal(serialComInt2) / 100))].Style.BackColor = Color.Green;
                }
            }
            else if (serialComInt3 == 800) //CONTROL BASE RADIO ERROR
            {
                GameSetup_datagridView.Rows[UnitRow].Cells[6 + Convert.ToInt16(Math.Floor(Convert.ToDecimal(serialComInt2) / 100))].Value = 0;
                GameSetup_datagridView.Rows[UnitRow].Cells[6 + Convert.ToInt16(Math.Floor(Convert.ToDecimal(serialComInt2) / 100))].Style.BackColor = Color.Red;
                //MessageBox.Show("CONTROL BASE RADIO ERROR: SENDING FAILED");
            }
            else if (serialComInt3 == 900) //CONTROL BASE RADIO ERROR
            {
                GameSetup_datagridView.Rows[UnitRow].Cells[6 + Convert.ToInt16(Math.Floor(Convert.ToDecimal(serialComInt2) / 100))].Value = 0;
                GameSetup_datagridView.Rows[UnitRow].Cells[6 + Convert.ToInt16(Math.Floor(Convert.ToDecimal(serialComInt2) / 100))].Style.BackColor = Color.Red;
                //MessageBox.Show("CONTROL BASE RADIO ERROR: NO REPLY FROM UNIT");
            }

            if (GameSetup_datagridView.Rows[UnitRow].Cells[6].Style.BackColor == Color.Red || GameSetup_datagridView.Rows[UnitRow].Cells[7].Style.BackColor == Color.Red)
            {
                GameSetup_datagridView.Rows[UnitRow].Cells[15].Style.BackColor = Color.Red;
            }
            else
            {
                GameSetup_datagridView.Rows[UnitRow].Cells[15].Style.BackColor = Color.Green;
            }

            StepStatus1_textBox.Text = "OK"; //INITIAL CONDITION, TO BE OVERIDDEN IF RED CELLS ARE FOUND IN "IF" LOOP
            StepStatus1_textBox.BackColor = System.Drawing.Color.Lime;
            for (int i = 0; i < GameSetup_datagridView.RowCount; i++)
            {
                if (GameSetup_datagridView.Rows[i].Cells[6].Style.BackColor == Color.Red || GameSetup_datagridView.Rows[i].Cells[7].Style.BackColor == Color.Red)
                { 
                    StepStatus1_textBox.Text = "NOK";
                    StepStatus1_textBox.BackColor = System.Drawing.Color.Red;
                }
            }
        }
        #endregion



        #region LOCK ACCESS TO TABS, AND RESTRICT TO BUTTON CONTROLS ONLY
        private void Game_tabControl_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Game_tabControl.SelectedIndex != tabSelected)
            {
                Game_tabControl.SelectedIndex = tabSelected;
            }
        }













        /*EXAMPLE OF CODE TO ADD WHEN SWITCHING TABS
        tabSelected = 1;  //Progress to second tab
        this.Invoke(new EventHandler(Game_tabControl_SelectedIndexChanged));
        */

        #endregion

        #region SEND GAME SETUP TO PLAYERS DEVICES - OLD
        //BUTTON - SEND GAME SETUP - BEGIN----------------------------------------------------------------------------------------
        /*
        private void GameSetup_SendSetup_button_Click(Object sender, EventArgs e) //BUTTON - SEND GAME SETUP
        {
            #region POPULATE DATAGRIDVIEW WITH PLAYERID'S
            for (int i = 0; i <= 99; i++)
            {
                if (GameSetup_datagridView.Rows[i].Cells[2].Value == null)
                {
                }
                else
                { 
                    for (int j = 0; j <= PlayerList.Count; j++)
                    {
                        if (GameSetup_datagridView.Rows[i].Cells[2].Value.ToString() == PlayerList[j].PlayerName)
                        {
                            GameSetup_datagridView.Rows[i].Cells[3].Value = PlayerList[j].PlayerID;
                        }
                            
                    }
                }      
            }
            #endregion


            try
            {
                if (SerialPortcomboBox.Text == "" || SerialRatecomboBox.Text == "") //Serial port settings not yet defined by user
                {
                    SerialRxtextBox.Text = "Must select Serial Port # and Rate before proceeding";
                }
                else //
                {
                    SerialRxtextBox.Text = ""; //Clear SerialRxtextBox
                    serialPort1.PortName = SerialPortcomboBox.Text; //Define port # per user selection
                    serialPort1.BaudRate = Convert.ToInt32(SerialRatecomboBox.Text); //Define port speed per user selection
                    serialPort1.Open(); //Now open the serial port

                    //TURN BUTTONS OFF (i.e. button1.Enabled = False)

                    //ADD CODE HERE FOR SENDING ALL GAME SETUP INFO

                    serialPort1.Close(); //Close the serial port when done
                }
            }
            catch (UnauthorizedAccessException) //Error response when disabled buttons are clicked during serial txrx
            {
                SerialRxtextBox.Text = "Unauthorized Access"; //Error response when disabled buttons are clicked during serial txrx
            }
        }
        */
        #endregion


    }
}
