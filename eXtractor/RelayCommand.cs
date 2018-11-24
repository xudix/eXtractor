using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace eXtractor
{
    public class RelayCommand : ICommand
    {
        public event EventHandler CanExecuteChanged;

        private readonly Action action;

        private readonly Func<bool> canExecute;

        public RelayCommand(Action action, Func<bool> canExecute)
        {
            this.action = action;
            this.canExecute = canExecute;
        }

        public RelayCommand(Action action)
        {
            this.action = action;
            this.canExecute = () => true;
        }

        public bool CanExecute(object parameter)
            => canExecute != null ? canExecute(): true;

        public void Execute(object parameter)
            => action();
    }
}
