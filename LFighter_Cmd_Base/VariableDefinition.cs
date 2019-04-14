using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LFighter_Cmd_Base
{
    static class VariableDefinition  //INITIALIZE VARIABLES FOR GAME SETTINGS OF MAIN FORM
    {

        //GAMETYPE
        public static List<string> GameType = new List<string>() { "Free for All", "Teams/Squads", "Survivor" };  //Initialize and populate GameType variable

        //GAMETIME
        public static List<int> GameTime1 = new List<int>() { 5, 10, 15, 20, 30, 45, 60, 90, 120 };  //Initialize and populate GameTime1 variable, for use in "Timed" games
        public static List<string> GameTime2 = new List<string>() { "N/A" };  //Initialize and populate GameTime2 variable, for use in "Elimination" games

        //PREGAMETIME
        public static List<int> PregameTime = new List<int>() { 1, 2, 3, 4, 5, 10, 15, 20 };  //Initialize and populate PregameTime variable

        //PLAYER ROSTER DEFAULTS FOR GameSetup_datagridView
        public static String defaultGunBattery = "----";
        public static String defaultVestBattery = "----";
        public static String defaultTeam = "1";
        public static String defaultHealth = "100";
        public static String defaultShield = "100";
        public static String defaultGunROF = "Semi-Auto";
        public static String defaultGunDamage = "10";
        public static String defaultAmmo = "100";
        public static Boolean defaultSave = true;
        public static String BattCheck = "Recheck";
        public static String BootUnit = "Reboot";
    }
}
