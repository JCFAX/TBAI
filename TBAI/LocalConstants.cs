using System;
using System.Collections.Generic;
using System.Text;
using SAPbobsCOM;

namespace TBAI
{
    public class LocalConstants
    {
        /// <summary>
        /// Clase con los identificadores de los formularios 
        /// </summary>
        public class FormIds
        {
            public const string frmAX_PEC = "AX_PEC";
        }

        /// <summary>
        /// Clase con las tablas de usuario necesarias para este addon
        /// </summary>
        public class UserTables
        {
        }

        /// <summary>
        /// Clase con los campos de las tablas de usuario
        /// </summary>
        public class UserFields
        {
        } 
        public class Colors
        {
            public const int BACKGROUND_GRAY = 15724527;
            public const int NORMAL_WHITE = 16777215;
            public const int BACKGROUND_RED = 11110925;
            public const int BACKGROUND_GREEN = 832430;
        }

        public const int TIMEOUT_AHORA = 1000000;
    }
}
