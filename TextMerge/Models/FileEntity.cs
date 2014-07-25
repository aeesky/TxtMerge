using TextMerge.Helpers;

namespace TextMerge.Models
{
    public class FileEntity : NotificationObject
    {
        #region Ctor
        public FileEntity()
        {
        }
        #endregion
        #region Property
        
        private int _no;
        /// <summary>
        /// 
        /// </summary>
        public int No
        {
            get { return _no; }
            set
            {
                if (_no != value)
                {
                    _no = value;
                    RaisePropertyChanged("No");
                }
            }
        }

        
        private string _ColumnName;
        /// <summary>
        /// 
        /// </summary>
        public string ColumnName
        {
            get { return _ColumnName; }
            set
            {
                if (_ColumnName != value)
                {
                    _ColumnName = value;
                    RaisePropertyChanged("ColumnName");
                }
            }
        }

        
        private string _filePath;
        /// <summary>
        /// 
        /// </summary>
        public string FilePath
        {
            get { return _filePath; }
            set
            {
                if (_filePath != value)
                {
                    _filePath = value;
                    RaisePropertyChanged("FilePath");
                }
            }
        }

        
        private bool _isDone;
        /// <summary>
        /// 
        /// </summary>
        public bool IsDone
        {
            get { return _isDone; }
            set
            {
                if (_isDone != value)
                {
                    _isDone = value;
                    RaisePropertyChanged("IsDone");
                }
            }
        }
        #endregion

    }
}
