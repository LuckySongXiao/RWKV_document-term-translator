using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using DocumentTranslator.Services.RwkvLocal;

namespace DocumentTranslator.Windows
{
    /// <summary>
    /// 模型选择对话框 - 让用户选择要下载的模型文件
    /// </summary>
    public partial class ModelSelectWindow : Window
    {
        /// <summary>
        /// 用户选中的文件列表
        /// </summary>
        public List<ModelScopeFile> SelectedFiles { get; private set; } = new List<ModelScopeFile>();

        private List<ModelFileItem> _items = new List<ModelFileItem>();

        public ModelSelectWindow(List<ModelScopeFile> files, string sourceUrl, string formatHint)
        {
            InitializeComponent();

            SourceText.Text = sourceUrl;

            // 转换为可绑定的项列表
            foreach (var f in files)
            {
                _items.Add(new ModelFileItem
                {
                    FileName = f.FileName,
                    Size = f.Size,
                    SizeBytes = f.SizeBytes,
                    DownloadUrl = f.DownloadUrl,
                    IsSelected = false
                });
            }

            FileListView.ItemsSource = _items;
            UpdateSummary();
        }

        private void SelectAllChanged(object sender, RoutedEventArgs e)
        {
            var isChecked = SelectAllCheckBox.IsChecked == true;
            foreach (var item in _items)
            {
                item.IsSelected = isChecked;
            }
            FileListView.Items.Refresh();
            UpdateSummary();
        }

        private void ItemSelectionChanged(object sender, RoutedEventArgs e)
        {
            UpdateSummary();
        }

        private void UpdateSummary()
        {
            var selectedCount = _items.Count(i => i.IsSelected);
            var selectedSize = _items.Where(i => i.IsSelected).Sum(i => i.SizeBytes);

            SummaryText.Text = $"已选择 {selectedCount}/{_items.Count} 个文件";
            TotalSizeText.Text = selectedSize > 0
                ? $"选中大小: {FormatSize(selectedSize)}"
                : "";

            // 同步全选框状态
            if (selectedCount == 0)
                SelectAllCheckBox.IsChecked = false;
            else if (selectedCount == _items.Count)
                SelectAllCheckBox.IsChecked = true;
            else
                SelectAllCheckBox.IsChecked = null; // 不确定状态
        }

        private void DownloadSelected(object sender, RoutedEventArgs e)
        {
            var selected = _items.Where(i => i.IsSelected).ToList();
            if (selected.Count == 0)
            {
                MessageBox.Show("请至少选择一个模型文件", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            SelectedFiles = selected.Select(i => new ModelScopeFile
            {
                FileName = i.FileName,
                DisplayName = i.FileName,
                Size = i.Size,
                SizeBytes = i.SizeBytes,
                DownloadUrl = i.DownloadUrl
            }).ToList();

            DialogResult = true;
        }

        private void CancelDownload(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        private static string FormatSize(long bytes)
        {
            if (bytes >= 1073741824) return $"{bytes / 1073741824.0:F2} GB";
            if (bytes >= 1048576) return $"{bytes / 1048576.0:F1} MB";
            if (bytes >= 1024) return $"{bytes / 1024.0:F1} KB";
            return $"{bytes} B";
        }
    }

    /// <summary>
    /// 可绑定的模型文件项（支持 IsSelected 双向绑定）
    /// </summary>
    public class ModelFileItem : INotifyPropertyChanged
    {
        private bool _isSelected;

        public string FileName { get; set; } = "";
        public string Size { get; set; } = "";
        public long SizeBytes { get; set; }
        public string DownloadUrl { get; set; } = "";

        public bool IsSelected
        {
            get => _isSelected;
            set { _isSelected = value; OnPropertyChanged(); }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
