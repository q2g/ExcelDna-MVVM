namespace ExcelDna_MVVM.Utils
{
    #region Usings
    using System;
    using System.Windows;
    using System.Windows.Data;
    #endregion

    public class BindingObject : DependencyObject, IDisposable
    {
        private bool _suppress;
        private Action<DependencyPropertyChangedEventArgs> _onChanged;
        public Action<DependencyPropertyChangedEventArgs> OnChanged
        {
            get => _onChanged; set
            {
                _onChanged = value;
                _onChanged?.Invoke(new DependencyPropertyChangedEventArgs(ValueProperty, null, CachedData));
            }
        }

        #region ctor
        public BindingObject(object source, string bindingPath, Action<DependencyPropertyChangedEventArgs> onChanged, bool supressFireOnBindingInit = true)
        {
            using (SuppressNotifications(supressFireOnBindingInit))
            {
                _onChanged = onChanged;

                binding = new Binding(bindingPath);
                binding.Source = source;
                binding.Mode = BindingMode.TwoWay;
                BindingOperations.SetBinding(this, ValueProperty, binding);
            }
        }
        #endregion

        private Binding binding;
        public object SourceObject
        {
            get
            {
                if (binding != null)
                    return binding.Source;
                return null;
            }
        }

        public object CachedData { get; private set; }


        public object Value
        {
            get { return GetValue(ValueProperty); }
            set { SetValue(ValueProperty, value); }
        }


        public static readonly DependencyProperty ValueProperty =
            DependencyProperty.Register("Value", typeof(object), typeof(BindingObject), new UIPropertyMetadata(null, ValueChangedCallback));

        private static void ValueChangedCallback(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var instance = (BindingObject)d;
            instance.CachedData = e.NewValue;
            Action<DependencyPropertyChangedEventArgs> onChanged = instance._onChanged;

            if (onChanged != null && !instance._suppress)
            {
                onChanged(e);
            }
        }

        public void Dispose()
        {
            BindingOperations.ClearBinding(this, ValueProperty);
            _onChanged = null;
        }

        public IDisposable SuppressNotifications(bool supress)
        {
            return new Supresser(this, supress);
        }

        private class Supresser : IDisposable
        {
            private BindingObject _bindingObject;

            public Supresser(BindingObject bindingObject, bool supress)
            {
                _bindingObject = bindingObject;
                _bindingObject._suppress = supress;
            }

            public void Dispose()
            {
                _bindingObject._suppress = false;
            }
        }
    }
}
