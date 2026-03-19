using System;
using System.Windows.Forms;
using ICSharpCode.SharpDevelop.Gui;

namespace ClarionAssistant
{
    /// <summary>
    /// Dockable pad for Clarion CLASS editing helper.
    /// </summary>
    public class ClassHelperPad : AbstractPadContent
    {
        private ClassHelperControl _control;

        public override Control Control
        {
            get
            {
                if (_control == null)
                {
                    _control = new ClassHelperControl();
                }
                return _control;
            }
        }

        public override void Dispose()
        {
            if (_control != null)
            {
                _control.Dispose();
                _control = null;
            }
            base.Dispose();
        }

        public override void RedrawContent()
        {
            _control?.Refresh();
        }
    }
}
