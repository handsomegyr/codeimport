using System.ComponentModel;

namespace ExcelApplication1
{
    public class Settings: INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged(PropertyChangedEventArgs e)
        {
            if (PropertyChanged != null)
                PropertyChanged(this, e);
        }

        private string prize_id;
        public string PrizeId
        {
            get { return prize_id; }
            set
            {
                prize_id = value;
                OnPropertyChanged(new PropertyChangedEventArgs("PrizeId"));
            }
        }


        private string start_time;
        public string StartTime
        {
            get { return start_time; }
            set
            {
                start_time = value;
                OnPropertyChanged(new PropertyChangedEventArgs("StartTime"));
            }
        }

        private string end_time;
        public string EndTime
        {
            get { return end_time; }
            set
            {
                end_time = value;
                OnPropertyChanged(new PropertyChangedEventArgs("EndTime"));
            }

        }

        private bool is_used;
        public bool IsUsed
        {
            get { return is_used; }
            set
            {
                is_used = value;
                OnPropertyChanged(new PropertyChangedEventArgs("IsUsed"));
            }
        }

        private string activity_id;
        public string ActivityId
        {
            get { return activity_id; }
            set
            {
                activity_id = value;
                OnPropertyChanged(new PropertyChangedEventArgs("ActivityId"));
            }
        }

        private string prjcode;
        public string Prjcode
        {
            get { return prjcode; }
            set
            {
                prjcode = value;
                OnPropertyChanged(new PropertyChangedEventArgs("Prjcode"));
            }
        }

        private string targetPath;
        public string TargetPath
        {
            get { return targetPath; }
            set
            {
                targetPath = value;
                OnPropertyChanged(new PropertyChangedEventArgs("TargetPath"));
            }
        }

        private int num;
        public int Quantity
        {
            get { return num; }
            set
            {
                num = value;
                OnPropertyChanged(new PropertyChangedEventArgs("Quantity"));
            }
        }

        private int code_length;
        public int CodeLength
        {
            get { return code_length; }
            set
            {
                code_length = value;
                OnPropertyChanged(new PropertyChangedEventArgs("CodeLength"));
            }

        }
    }
}
