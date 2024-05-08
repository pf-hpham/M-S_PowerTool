using System;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;

namespace ColorChange
{
    public partial class Page1 : Window
    {
        public event EventHandler<CursorPositionEventArgs> CursorPositionChanged;

        private Line horizontalLine;
        private Line verticalLine;

        public Page1()
        {
            InitializeComponent();

            canvas.Width = 1920;
            canvas.Height = 1080;

            canvas.MouseMove += Canvas_MouseMove;
            canvas.MouseDown += Canvas_MouseDown;

            horizontalLine = new Line();
            verticalLine = new Line();

            horizontalLine.Stroke = Brushes.Red;
            horizontalLine.StrokeThickness = 1;
            DoubleCollection dashes_x = new DoubleCollection { 2, 2 };
            horizontalLine.StrokeDashArray = dashes_x;
            horizontalLine.X1 = 0;
            horizontalLine.X2 = 1920;
            horizontalLine.Y1 = 0;
            horizontalLine.Y2 = 0;

            verticalLine.Stroke = Brushes.Red;
            verticalLine.StrokeThickness = 1;
            DoubleCollection dashes_y = new DoubleCollection { 2, 2 };
            verticalLine.StrokeDashArray = dashes_y;
            verticalLine.X1 = 0;
            verticalLine.X2 = 0;
            verticalLine.Y1 = 0;
            verticalLine.Y2 = 1080;

            canvas.MouseMove += Canvas_MouseMove;
            canvas.Children.Add(horizontalLine);
            canvas.Children.Add(verticalLine);
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            double screenWidth = SystemParameters.PrimaryScreenWidth;
            double screenHeight = SystemParameters.PrimaryScreenHeight;

            Width = screenWidth;
            Height = screenHeight;
        }

        private void Canvas_MouseMove(object sender, MouseEventArgs e)
        {
            Point cursorPosition = e.GetPosition(canvas);

            horizontalLine.Y1 = cursorPosition.Y;
            horizontalLine.Y2 = cursorPosition.Y;

            verticalLine.X1 = cursorPosition.X;
            verticalLine.X2 = cursorPosition.X;

            pos.Text = $"X: {cursorPosition.X}, Y: {cursorPosition.Y}";
        }

        private void Canvas_MouseDown(object sender, MouseButtonEventArgs e)
        {
            Point clickPosition = e.GetPosition(canvas);
            OnCursorPositionChanged(clickPosition);
            Close();
        }

        protected virtual void OnCursorPositionChanged(Point cursorPosition)
        {
            CursorPositionChanged?.Invoke(this, new CursorPositionEventArgs(cursorPosition));
        }
    }

    public class CursorPositionEventArgs : EventArgs
    {
        public Point CursorPosition { get; }

        public CursorPositionEventArgs(Point cursorPosition)
        {
            CursorPosition = cursorPosition;
        }
    }
}
