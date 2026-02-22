using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportManager.Addon.Core
{
    public sealed class MenuManager
    {
        private readonly Application _app;

        public MenuManager(Application app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
        }

        public void EnsurePopupWithEntry(string parentMenuId, string popupMenuId, string popupCaption, string childMenuId, string childCaption)
        {
            if (string.IsNullOrWhiteSpace(parentMenuId))
                throw new ArgumentException("parentMenuId es requerido.", nameof(parentMenuId));

            if (string.IsNullOrWhiteSpace(popupMenuId))
                throw new ArgumentException("popupMenuId es requerido.", nameof(popupMenuId));

            if (string.IsNullOrWhiteSpace(childMenuId))
                throw new ArgumentException("childMenuId es requerido.", nameof(childMenuId));

            if (!_app.Menus.Exists(popupMenuId))
            {
                var popupParams = (MenuCreationParams)_app.CreateObject(BoCreatableObjectType.cot_MenuCreationParams);
                popupParams.Type = BoMenuType.mt_POPUP;
                popupParams.UniqueID = popupMenuId;
                popupParams.String = popupCaption;
                popupParams.Image = "";
                _app.Menus.Item(parentMenuId).SubMenus.AddEx(popupParams);
            }

            if (!_app.Menus.Exists(childMenuId))
            {
                var childParams = (MenuCreationParams)_app.CreateObject(BoCreatableObjectType.cot_MenuCreationParams);
                childParams.Type = BoMenuType.mt_STRING;
                childParams.UniqueID = childMenuId;
                childParams.String = childCaption;
                childParams.Image = "";
                _app.Menus.Item(popupMenuId).SubMenus.AddEx(childParams);
            }
        }
    }

}
