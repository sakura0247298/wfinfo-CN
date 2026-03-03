using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows.Data;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;

namespace WFInfo
{
    public class RelicsViewModel : INPC
    {
        public static RelicsViewModel Instance { get; }= new RelicsViewModel();
        public RelicsViewModel()
        {
            _relicTreeItems = new List<TreeNode>();
            RelicsItemsView = new ListCollectionView(_relicTreeItems);
            ExpandAllCommand = new SimpleCommand(() => ExpandOrCollapseAll(true));
            CollapseAllCommand = new SimpleCommand(() => ExpandOrCollapseAll(false));
        }

        private bool _initialized = false;
        private string _filterText = "";
        private int searchTimerDurationMS = 500;
        private bool _showAllRelics;
        private readonly List<TreeNode> _relicTreeItems;
        private int _sortBoxSelectedIndex;
        private bool _hideVaulted = true;
        private readonly List<TreeNode> _rawRelicNodes = new List<TreeNode>();

        public static Timer searchTimer = new Timer();

        private void StartSearchReapplyTimer()
        {
            if (searchTimer.Enabled)
            {
                searchTimer.Stop();
            }

            searchTimer.Interval = searchTimerDurationMS;
            searchTimer.Enabled = true;
            searchTimer.Tick += (s, e) =>
            {
                searchTimer.Enabled = false;
                searchTimer.Stop();
                ReapplyFilters();
            };
        }

        public string FilterText
        {
            get => _filterText;
            set
            {
                this.SetField(ref _filterText, value);
                StartSearchReapplyTimer();
                RaisePropertyChanged(nameof(IsFilterEmpty));
            }
        }
        public bool IsFilterEmpty => string.IsNullOrEmpty(FilterText);

        public SimpleCommand ExpandAllCommand { get; }
        public SimpleCommand CollapseAllCommand { get; }

        private void ExpandOrCollapseAll(bool expand)
        {
            foreach (TreeNode era in _rawRelicNodes)
                era.ChangeExpandedTo(expand);
        }
        
        public bool ShowAllRelics
        {
            get => _showAllRelics;
            set
            {
                foreach (TreeNode era in _rawRelicNodes)
                foreach (TreeNode relic in era.Children)
                    relic.topLevel = value;
                SetField(ref _showAllRelics, value);
                
                _relicTreeItems.Clear();
                RefreshVisibleRelics();
                RaisePropertyChanged(nameof(ShowAllRelicsText));
            }
        }
        public string ShowAllRelicsText => ShowAllRelics ? "所有遗物" : "遗物纪元";

        public bool HideVaulted
        {
            get => _hideVaulted;
            set
            { 
                SetField(ref _hideVaulted, value);
                ReapplyFilters();
            }
        }

        public ICollectionView RelicsItemsView { get; }

        public int SortBoxSelectedIndex
        {
            get => _sortBoxSelectedIndex;
            set
            {
                SetField(ref _sortBoxSelectedIndex, value);
                SortBoxChanged();
            }
        }
        public void SortBoxChanged()
        {
            // 0 - Name
            // 1 - Average intact plat
            // 2 - Average radiant plat
            // 3 - Difference (radiant-intact)
        
            foreach (TreeNode era in _rawRelicNodes)
            {
                era.Sort(SortBoxSelectedIndex);
                era.RecolorChildren();
            }
            if (ShowAllRelics)
            {
                RelicsItemsView.SortDescriptions.Clear();
                //TODO:
                //_relicTreeItems.IsLiveSorting = true;
                switch (SortBoxSelectedIndex)
                {
                    case 1:
                        RelicsItemsView.SortDescriptions.Add(new System.ComponentModel.SortDescription("Intact_Val", System.ComponentModel.ListSortDirection.Descending));
                        break;
                    case 2:
                        RelicsItemsView.SortDescriptions.Add(new System.ComponentModel.SortDescription("Radiant_Val", System.ComponentModel.ListSortDirection.Descending));
                        break;
                    case 3:
                        RelicsItemsView.SortDescriptions.Add(new System.ComponentModel.SortDescription("Bonus_Val", System.ComponentModel.ListSortDirection.Descending));
                        break;
                    default:
                        RelicsItemsView.SortDescriptions.Add(new System.ComponentModel.SortDescription("Name_Sort", System.ComponentModel.ListSortDirection.Ascending));
                        break;
                }

                bool i = false;
                foreach (TreeNode relic in _relicTreeItems)
                {
                    i = !i;
                    if (i)
                        relic.Background_Color = TreeNode.BACK_D_BRUSH;
                    else
                        relic.Background_Color = TreeNode.BACK_U_BRUSH;
                }
            }
        }

        public void RefreshVisibleRelics()
        {
            int index = 0;
            if (ShowAllRelics)
            {
                List<TreeNode> activeNodes = new List<TreeNode>();
                foreach (TreeNode era in _rawRelicNodes)
                foreach (TreeNode relic in era.ChildrenFiltered)
                    activeNodes.Add(relic);


                for (index = 0; index < _relicTreeItems.Count;)
                {
                    TreeNode relic = (TreeNode)_relicTreeItems.ElementAt(index);
                    if (!activeNodes.Contains(relic))
                        _relicTreeItems.RemoveAt(index);
                    else
                    {
                        activeNodes.Remove(relic);
                        index++;
                    }
                }

                foreach (TreeNode relic in activeNodes)
                    _relicTreeItems.Add(relic);

                SortBoxChanged();
            }
            else
            {
                foreach (TreeNode era in _rawRelicNodes)
                {
                    int curr = _relicTreeItems.IndexOf(era);
                    if (era.ChildrenFiltered.Count == 0)
                    {
                        if (curr != -1)
                            _relicTreeItems.RemoveAt(curr);
                    }
                    else
                    {
                        if (curr == -1)
                            _relicTreeItems.Insert(index, era);

                        index++;
                    }
                    era.RecolorChildren();
                }
            }
            RelicsItemsView.Refresh();
        }

        public void ReapplyFilters()
        {
        
            foreach (TreeNode era in _rawRelicNodes)
            {
                era.ResetFilter();
                if(HideVaulted)
                    era.FilterOutVaulted(true);
                if(!string.IsNullOrEmpty(FilterText))
                {
                    var searchText = FilterText.Split(' ');
                    era.FilterSearchText(searchText, false, true);
                }
            }
            RefreshVisibleRelics();
        }

        public void InitializeTree()
        {
            if (_initialized)
            {
                return;
            }
            TreeNode lith = new TreeNode("古纪", "", false, 0) { Era = "Lith" };
            TreeNode meso = new TreeNode("前纪", "", false, 0) { Era = "Meso" };
            TreeNode neo = new TreeNode("中纪", "", false, 0) { Era = "Neo" };
            TreeNode axi = new TreeNode("后纪", "", false, 0) { Era = "Axi" };
            TreeNode vanguard = new TreeNode("先锋", "", false, 0) { Era = "Vanguard" };
            _rawRelicNodes.AddRange(new[] { lith, meso, neo, axi, vanguard });
            int eraNum = 0;
            foreach (TreeNode head in _rawRelicNodes)
            {
                head.SortNum = eraNum++;
                foreach (JProperty prop in Main.dataBase.relicData[head.Era])
                {
                    JObject primeItems = (JObject)Main.dataBase.relicData[head.Era][prop.Name];
                    string vaulted = primeItems["vaulted"].ToObject<bool>() ? "已入库" : "";
                    TreeNode relic = new TreeNode(prop.Name, vaulted, false, 0);
                    relic.Era = head.Name;
                    foreach (KeyValuePair<string, JToken> kvp in primeItems)
                    {
                        if (kvp.Key != "vaulted" && Main.dataBase.marketData.TryGetValue(kvp.Value.ToString(), out JToken marketValues))
                        {
                            string partName = kvp.Value.ToString();
                            string localePartName = Main.dataBase.GetLocaleNameData(partName);
                            if (!string.IsNullOrEmpty(localePartName))
                                partName = localePartName;
                            TreeNode part = new TreeNode(partName, "", false, 0);
                            part.SetPartText(marketValues["plat"].ToObject<double>(), marketValues["ducats"].ToObject<int>(), kvp.Key);                           
                            relic.AddChild(part);
                        }
                    }
                                
                    relic.SetRelicText();
                    head.AddChild(relic);
            
                    //groupedByAll.Items.Add(relic);
                    //Search.Items.Add(relic);
                }
            
                head.SetEraText();   
                head.ResetFilter();
                head.FilterOutVaulted();
                head.RecolorChildren();
            }
            RefreshVisibleRelics();
            SortBoxChanged();
            _initialized = true;
        }
    }
}