using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LTag_LaunchPad
{
    public class PlayerData //Sets up object variable for importing historical player data
    {
        //INITIALIZE FIELDS FOR HISTORICAL DATA - BEGIN----------------------------------------------------------------------------------------------
        public string PlayerName { get; set; } //Variable from Excel file with Player Historical Data
        public double PlayerID { get; set; } //Variable from Excel file with Player Historical Data
        public double WinPerc { get; set; } //Variable from Excel file with Player Historical Data
        public double AccuracyPerc { get; set; } //Variable from Excel file with Player Historical Data
        public double HitsperGame { get; set; } //Variable from Excel file with Player Historical Data
        public double HeadshotsperGame { get; set; } //Variable from Excel file with Player Historical Data
        public double DamageperGame { get; set; } //Variable from Excel file with Player Historical Data
        public double EliminationsperGame { get; set; } //Variable from Excel file with Player Historical Data
        public double RevivesperGame { get; set; } //Variable from Excel file with Player Historical Data
        public double DamageTakenperGame { get; set; } //Variable from Excel file with Player Historical Data
        //INITIALIZE FIELDS FOR HISTORICAL DATA - END^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    }
}
