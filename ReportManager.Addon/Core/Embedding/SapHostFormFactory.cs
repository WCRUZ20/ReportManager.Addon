using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ReportManager.Addon.Core.Embedding
{
    internal static class SapHostFormFactory
    {
        public static Form TryGetOpenForm(Application app, string uid)
        {
            try { return app.Forms.Item(uid); }
            catch { return null; }
        }

        public static Form CreateOrFocusHostForm(Application app, string uid, string title, int width, int height, int left = 300, int top = 80)
        {
            var existing = TryGetOpenForm(app, uid);
            if (existing != null)
            {
                existing.Visible = true;
                existing.Select();
                return existing;
            }

            var form = app.Forms.Add(uid, BoFormTypes.ft_Fixed);
            form.Title = title;
            form.Left = left;
            form.Top = top;
            form.Width = width;
            form.Height = height;
            form.Visible = true;
            form.Select();
            return form;
        }

        public static IntPtr GetHostFormHwnd(Application app, Form hostForm, IntPtr sapMainHwnd)
        {
            // Se basa en el Title (como tu guía VB). Reintenta porque SAP puede tardar.
            for (int i = 0; i < 12; i++)
            {
                var hwnd = SapWindowHandles.FindChildWindowByExactTitle(sapMainHwnd, hostForm.Title);
                if (hwnd != IntPtr.Zero) return hwnd;
                Thread.Sleep(50);
            }

            return IntPtr.Zero;
        }
    }
}
