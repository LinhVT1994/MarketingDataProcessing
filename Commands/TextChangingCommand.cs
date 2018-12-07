using MarketingDataProcessing.Models;
using MarketingDataProcessing.Utilities;
using MarketingDataProcessing.ViewModels;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace MarketingDataProcessing.Commands
{
    class TextChangingCommand : ICommand
    {
        SearchViewModel _ViewModel;

        public TextChangingCommand(SearchViewModel viewModel)
        {
            _ViewModel = viewModel;
        }

        public bool IsDisplayNotFoundMessage { get; private set; }
        public bool IsDisplaySuggestion { get; private set; }

        public event EventHandler CanExecuteChanged
        {
            add
            {
                CommandManager.RequerySuggested += value;
            }
            remove
            {
                CommandManager.RequerySuggested -= value;
            }
        }

        public bool CanExecute(object parameter)
        {
            return true;
        }
        public void Execute(object parameter)
        {
            _ViewModel.SearchForSuggestionResult(parameter);
        }
    }
}
