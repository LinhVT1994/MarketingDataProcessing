using MarketingDataProcessing.Commands;
using MarketingDataProcessing.Models;
using MarketingDataProcessing.Utilities;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace MarketingDataProcessing.ViewModels
{
    class SearchViewModel:BaseMVVM
    {
        private bool _IsDisplaySuggestion = false;
        private bool _IsDisplayNotFoundMessage = false;
        private ObservableCollection<Searching> _SearchedResults;
        public string NotFoundMessage
        {
            get
            {
                return "探したいデータがありません。申し訳ございません";
            }
        }
        public ICommand OnMouseDownOnResultItem
        {
            get;
            set;
        }
        public ICommand ItemSelectedCommand
        {
            get;
            set;
        }
        public ICommand OnEnterKeyUpEvent;
        public ICommand OnTextChangingEvent
        {
            get;
            set;
        }
        public ICommand OnHistoryItemEnterCommand
        {
            get;
            set;
        }
        public SearchViewModel()
        {
            OnTextChangingEvent = new TextChangingCommand(this);
        }
      
        public ObservableCollection<Searching> SuggestionsValues
        {
            get
            {
                return _SearchedResults;
            }
            set
            {
                _SearchedResults = value;
                RaisePropertyChanged(nameof(SuggestionsValues));
            }
        }
        public bool IsDisplayNotFoundMessage
        {

            get
            {
                return _IsDisplayNotFoundMessage;
            }
            set
            {
                _IsDisplayNotFoundMessage = value;
                RaisePropertyChanged(nameof(IsDisplayNotFoundMessage));
            }
        }
        public bool IsDisplaySuggestion
        {

            get
            {
                return _IsDisplaySuggestion;
            }
            set
            {
                _IsDisplaySuggestion = value;
                RaisePropertyChanged(nameof(IsDisplaySuggestion));
            }
        }

        internal void SearchForSuggestionResult(object parameter)
        {

            SqlDataAccess sqlDataAccess = new SqlDataAccess();
            string sqlString = @"select * from seaching where fulltext like @para limit 50";
            List<SqlParameter> paraneters = new List<SqlParameter>
            {
                new SqlParameter("para","%"+parameter.ToString()+"%")
            };
            DataTable results = sqlDataAccess.ExecuteSelectQuery(sqlString, paraneters.ToArray());

            if (results.Count == 0)
            {
                IsDisplaySuggestion = false;
                IsDisplayNotFoundMessage = true;
            }
            else
            {
                IsDisplayNotFoundMessage = false;
                IsDisplaySuggestion = true;
                SuggestionsValues = new ObservableCollection<Searching>();
                int count = 0;
                foreach (var result in results.GetAllRecords())
                {
                    var record = Searching.CreateNew(result);
                    record.Id = ++count;
                    SuggestionsValues.Add(record);
                }
            }

        }
    }
}
