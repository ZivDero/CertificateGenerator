using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace CertificateGenerator.ViewModel
{
    public class ColorToBrushConverter : IValueConverter
    { 
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is Color color)
                return new SolidColorBrush(color);
            return new SolidColorBrush(Colors.Black);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is SolidColorBrush brush)
                return brush.Color;
            return Colors.Black;
        }
    }

    public class LeftAlignmentConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is MainViewModel.TextAlignment alignment)
                return alignment == MainViewModel.TextAlignment.Left;
            return false;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool alignment)
                if (alignment)
                    return MainViewModel.TextAlignment.Left;
            return MainViewModel.TextAlignment.Center;
        }
    }

    public class CenterAlignmentConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is MainViewModel.TextAlignment alignment)
                return alignment == MainViewModel.TextAlignment.Center;
            return false;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool alignment)
                if (alignment)
                    return MainViewModel.TextAlignment.Center;
            return MainViewModel.TextAlignment.Center;
        }
    }

    public class RightAlignmentConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is MainViewModel.TextAlignment alignment)
                return alignment == MainViewModel.TextAlignment.Right;
            return false;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool alignment)
                if (alignment)
                    return MainViewModel.TextAlignment.Right;
            return MainViewModel.TextAlignment.Center;
        }
    }
}
