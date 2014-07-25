using System;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Windows.Forms;
using System.Windows.Input;
using TextMerge.Helpers;
using TextMerge.Models;
using TextMerge.Properties;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;

namespace TextMerge.ViewModels
{
    public class MainWindowViewModel : BaseViewModel
    {
        #region Properties

        private string _sourcePath;

        /// <summary>
        /// </summary>
        public string SourcePath
        {
            get { return _sourcePath; }
            set
            {
                if (_sourcePath != value)
                {
                    _sourcePath = value;
                    RaisePropertyChanged("SourcePath");
                }
            }
        }


        private string _destFile;

        /// <summary>
        /// </summary>
        public string DestFile
        {
            get { return _destFile; }
            set
            {
                if (_destFile != value)
                {
                    _destFile = value;
                    RaisePropertyChanged("DestFile");
                }
            }
        }

        private readonly ExcelHelper _Excel;

        #endregion

        #region FilesCollection

        private ObservableCollection<FileEntity> _FilesCollection;

        public ObservableCollection<FileEntity> FilesCollection
        {
            get { return _FilesCollection; }
            set
            {
                if (_FilesCollection != value)
                {
                    _FilesCollection = value;
                    RaisePropertyChanged(() => FilesCollection);
                }
            }
        }

        #endregion

        #region Commands

        public ICommand SelectPathCommand
        {
            get { return new DelegateCommand(SelectPath); }
        }

        public ICommand SelectFileCommand
        {
            get { return new DelegateCommand(SelectFile); }
        }

        public ICommand MergeCommand
        {
            get { return new DelegateCommand(MergeFile, CanExecuteMerge); }
        }

        #endregion

        #region Ctor

        public MainWindowViewModel()
        {
            _Excel = new ExcelHelper();
        }

        #endregion

        #region Command Handlers

        private void SelectPath()
        {
            var sfDialog = new FolderBrowserDialog();
            if (sfDialog.ShowDialog() == DialogResult.OK)
            {
                SourcePath = sfDialog.SelectedPath;
                FilesCollection = new ObservableCollection<FileEntity>();
                int no = 1;
                foreach (string file in DiretoryHelper.GetFiles(SourcePath))
                {
                    var entity = new FileEntity
                    {
                        No = no++,
                        ColumnName = Path.GetDirectoryName(file),
                        FilePath = file,
                        IsDone = false
                    };
                    FilesCollection.Add(entity);
                }
                RaisePropertyChanged(() => FilesCollection);
            }
        }

        private void SelectFile()
        {
            var opfDialog = new OpenFileDialog {Filter = "Excel文件(*.xls)|*.xls"};
            bool? ret = opfDialog.ShowDialog();
            if (ret.HasValue && ret.Value)
            {
                DestFile = opfDialog.FileName;
            }
        }

        private void MergeFile()
        {
            if (string.IsNullOrEmpty(DestFile))
            {
                DestFile = Path.Combine(Directory.GetCurrentDirectory(), DateTime.Now.ToString("yyyy-M-d") + ".xls");
                _Excel.Create(DestFile);
            }
            else if (!_Excel.Open(DestFile))
            {
                MessageBox.Show(Resources.MainWindowViewModel_MergeFile_open_failed + DestFile);
                return;
            }

            //标记是否加入第一列数据
            bool isall = true;

            //记录插入列的位置
            int count = 0;
            foreach (FileEntity file in FilesCollection)
            {
                try
                {
#if RELEASE
                    #region 异步方法
                    MergeDelegate merge = new MergeDelegate(MergeOne);
                    var result = merge.BeginInvoke(file, isall, count,
                        (p) =>
                        {
                            isall = false;
                            file.IsDone = true;
                            RaisePropertyChanged(() => FilesCollection);
                        }, null);
                    int c = merge.EndInvoke(result);
                    count += c;
                    file.IsDone = true;
                    RaisePropertyChanged(() => FilesCollection); 
                    #endregion
#else

                    #region 同步方法
                    DataTable dt = TextHelper.GetTextData(file.FilePath, isall);
                    isall = false;
                    _Excel.ImportFromTable(dt, 0, count, 0);
                    count += dt.Columns.Count;
                    file.IsDone = true;
                    RaisePropertyChanged(() => FilesCollection);
                    #endregion
#endif
                }
                catch (Exception exception)
                {
                    Console.WriteLine(exception.Message);
                }
            }
            _Excel.Save();
            MessageBox.Show(Resources.MainWindowViewModel_MergeFile_done + DestFile);
        }

        private delegate int MergeDelegate(FileEntity file, bool isall, int count);

        private int MergeOne(FileEntity file, bool isall, int count)
        {
            DataTable dt = TextHelper.GetTextData(file.FilePath, isall);
            try
            {
                _Excel.ImportFromTable(dt, 0, count, 0);
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.Message);
            }
            return dt.Columns.Count;
        }

        private bool CanExecuteMerge()
        {
            return !string.IsNullOrEmpty(SourcePath);
        }

        #endregion
    }
}