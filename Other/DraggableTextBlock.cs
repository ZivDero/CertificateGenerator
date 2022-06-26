using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace CertificateGenerator.Other
{
    public class DraggableTextBlock : TextBlock
    {
        private Point offset;
        private bool isBeingDragged;

        public static readonly DependencyProperty RealWidthProperty = DependencyProperty.Register(
            "RealWidth", typeof(double), typeof(DraggableTextBlock), new PropertyMetadata(default(double)));

        public double RealWidth
        {
            get => (double) GetValue(RealWidthProperty);
            set => SetValue(RealWidthProperty, value);
        }

        public static readonly DependencyProperty RealHeightProperty = DependencyProperty.Register(
            "RealHeight", typeof(double), typeof(DraggableTextBlock), new PropertyMetadata(default(double)));

        public double RealHeight
        {
            get => (double)GetValue(RealHeightProperty);
            set => SetValue(RealHeightProperty, value);
        }

        public static readonly DependencyProperty PositionProperty = DependencyProperty.Register(
            "Position", typeof(Point), typeof(DraggableTextBlock), new PropertyMetadata(default(Point)));

        public Point Position
        {
            get => (Point) GetValue(PositionProperty);
            set => SetValue(PositionProperty, value);
        }

        public DraggableTextBlock()
        {
            Loaded += OnLoaded;
            PreviewMouseLeftButtonUp += OnPreviewMouseLeftButtonUp;
            PreviewMouseLeftButtonDown += OnPreviewMouseLeftButtonDown;
            PreviewMouseMove += OnPreviewMouseMove;
            SizeChanged += OnSizeChanged;
            MouseLeave += OnMouseLeave;

            RealWidth = ActualWidth;
            RealHeight = ActualHeight;
        }

        private void OnPreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (isBeingDragged)
                isBeingDragged = false;
            e.Handled = true;
        }

        private void OnPreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (!isBeingDragged)
            {
                isBeingDragged = true;
                offset = new Point(Mouse.GetPosition(Application.Current.MainWindow).X - Margin.Left, Mouse.GetPosition(Application.Current.MainWindow).Y - Margin.Top);
            }
            e.Handled = true;
        }

        private void OnPreviewMouseMove(object sender, MouseEventArgs e)
        {
  
           if (isBeingDragged)
           {
               Point mousePoint = Mouse.GetPosition(Application.Current.MainWindow);
               double x = mousePoint.X - offset.X > 0 ? mousePoint.X - offset.X : 0;
               double y = mousePoint.Y - offset.Y > 0 ? mousePoint.Y - offset.Y : 0;
               Margin = new Thickness(x, y, 0, 0);
               Position = TransformToAncestor(Window.GetWindow(this)).Transform(new Point(0, 0));
           }

           e.Handled = true;
        }

        private void OnPreviewKeyDownEvent(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
                isBeingDragged = false;
        }

        private void OnSizeChanged(object sender, SizeChangedEventArgs e)
        {
            RealWidth = ActualWidth;
            RealHeight = ActualHeight;
            e.Handled = true;
        }

        private void OnMouseLeave(object sender, MouseEventArgs e)
        {
            isBeingDragged = false;
            e.Handled = true;
        }

        private void OnLoaded(object sender, EventArgs e)
        {
            Window.GetWindow(this).PreviewKeyDown += OnPreviewKeyDownEvent;
            Position = TransformToAncestor(Window.GetWindow(this)).Transform(new Point(0, 0));
        }
    }
}
