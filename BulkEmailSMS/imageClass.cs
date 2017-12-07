using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using System.IO;

namespace BulkEmailSMS
{

    public class DataGridViewImageButtonDeleteColumn : DataGridViewButtonColumn
    {

        public DataGridViewImageButtonDeleteColumn()
        {
            this.CellTemplate = new DataGridViewImageButtonDeleteCell();
            this.Width = 22;
            this.Resizable = DataGridViewTriState.False;

        }
    }
    public class DataGridViewImageButtonDeleteCell : DataGridViewImageButtonCell
    {

        public override void LoadImages()
        {

            try
            {


                string appFolderPath = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);

                string resourcesFolderPath = Path.Combine(
                Directory.GetParent(appFolderPath).Parent.FullName, @"Resources\");

                _buttonImageDisabled = Image.FromFile(resourcesFolderPath + "Delete.png");


                ButtonState = PushButtonState.Disabled;




            }
            catch (Exception ex)
            {


            }



        }
    }



    //>>
    public class DataGridViewImageButtonEditColumn : DataGridViewButtonColumn
    {
        public DataGridViewImageButtonEditColumn()
        {
            this.CellTemplate = new DataGridViewImageButtonEditCell();
            this.Width = 22;
            this.Resizable = DataGridViewTriState.False;

        }
    }
    public class DataGridViewImageButtonEditCell : DataGridViewImageButtonCell
    {
        public override void LoadImages()
        {

            try
            {
                string appFolderPath = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);

                string resourcesFolderPath = Path.Combine(
                Directory.GetParent(appFolderPath).Parent.FullName, @"Resources\");

                _buttonImageNormal = Image.FromFile(resourcesFolderPath + @"edit.png");
                ButtonState = PushButtonState.Default;




            }
            catch (Exception ex)
            {


            }



        }
    }

    //>>


    //>>>>


    public class DataGridViewImageButtonProfBtnColumn : DataGridViewButtonColumn
    {
        public DataGridViewImageButtonProfBtnColumn()
        {
            this.CellTemplate = new DataGridViewImageButtonProfBtnCell();
            this.Width = 22;
            this.Resizable = DataGridViewTriState.False;

        }
    }
    public class DataGridViewImageButtonProfBtnCell : DataGridViewImageButtonCell
    {
        public override void LoadImages()
        {

            try
            {
                string appFolderPath = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);

                string resourcesFolderPath = Path.Combine(
                Directory.GetParent(appFolderPath).Parent.FullName, @"Resources\");


                _buttonImageHot = Image.FromFile(resourcesFolderPath + @"prof1.png");

                ButtonState = PushButtonState.Hot;




            }
            catch (Exception ex)
            {


            }



        }
    }




    //>>>>






    public abstract class DataGridViewImageButtonCell : DataGridViewButtonCell
    {
        private bool _enabled;                // Is the button enabled
        private PushButtonState _buttonState; // What is the button state
        protected Image _buttonImageHot;      // The hot image
        protected Image _buttonImageNormal;   // The normal image
        protected Image _buttonImageDisabled; // The disabled image
        private int _buttonImageOffset;       // The amount of offset or border around the image

        protected DataGridViewImageButtonCell()
        {
            // In my project, buttons are disabled by default
            _enabled = false;
            _buttonState = PushButtonState.Disabled;

            // Changing this value affects the appearance of the image on the button.
            _buttonImageOffset = 2;

            // Call the routine to load the images specific to a column.
            LoadImages();
        }

        // Button Enabled Property
        public bool Enabled
        {
            get
            {
                return _enabled;
            }

            set
            {
                _enabled = value;
                _buttonState = value ? PushButtonState.Normal : PushButtonState.Disabled;
            }
        }

        // PushButton State Property
        public PushButtonState ButtonState
        {
            get { return _buttonState; }
            set { _buttonState = value; }
        }

        // Image Property
        // Returns the correct image based on the control's state.
        public Image ButtonImage
        {
            get
            {
                switch (_buttonState)
                {
                    case PushButtonState.Disabled:
                        return _buttonImageDisabled;

                    case PushButtonState.Hot:
                        return _buttonImageHot;

                    case PushButtonState.Normal:
                        return _buttonImageNormal;

                    case PushButtonState.Pressed:
                        return _buttonImageNormal;

                    case PushButtonState.Default:
                        return _buttonImageNormal;

                    default:
                        return _buttonImageNormal;
                }
            }
        }

        protected override void Paint(Graphics graphics,
            Rectangle clipBounds, Rectangle cellBounds, int rowIndex,
            DataGridViewElementStates elementState, object value,
            object formattedValue, string errorText,
            DataGridViewCellStyle cellStyle,
            DataGridViewAdvancedBorderStyle advancedBorderStyle,
            DataGridViewPaintParts paintParts)
        {
            //base.Paint(graphics, clipBounds, cellBounds, rowIndex, elementState, value, formattedValue, errorText, cellStyle, advancedBorderStyle, paintParts);

            // Draw the cell background, if specified.
            if ((paintParts & DataGridViewPaintParts.Background) ==
                DataGridViewPaintParts.Background)
            {
                SolidBrush cellBackground =
                    new SolidBrush(cellStyle.BackColor);
                graphics.FillRectangle(cellBackground, cellBounds);
                cellBackground.Dispose();
            }

            // Draw the cell borders, if specified.
            if ((paintParts & DataGridViewPaintParts.Border) ==
                DataGridViewPaintParts.Border)
            {
                PaintBorder(graphics, clipBounds, cellBounds, cellStyle,
                    advancedBorderStyle);
            }

            // Calculate the area in which to draw the button.
            // Adjusting the following algorithm and values affects
            // how the image will appear on the button.
            Rectangle buttonArea = cellBounds;

            Rectangle buttonAdjustment =
                BorderWidths(advancedBorderStyle);

            buttonArea.X += buttonAdjustment.X;
            buttonArea.Y += buttonAdjustment.Y;
            buttonArea.Height -= buttonAdjustment.Height;
            buttonArea.Width -= buttonAdjustment.Width;

            Rectangle imageArea = new Rectangle(
                buttonArea.X + _buttonImageOffset,
                buttonArea.Y + _buttonImageOffset,
                16,
                16);

            ButtonRenderer.DrawButton(graphics, buttonArea, ButtonImage, imageArea, false, ButtonState);
        }


        public abstract void LoadImages();
    }



}
