using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Autodiscover;
using Microsoft.Exchange.WebServices.Data;

namespace Microsoft.Samples.MigrationTools
{

    public interface IMigrateItem
    {
        void CreateItem();
        void SetProperty();
        void SetProperties();

        void SaveItem();
    }

    class MAPI2EWS : IMigrateItem
    {
        public void CreateItem()
        {

        }

        void IMigrateItem.SetProperty()
        {
            throw new NotImplementedException();
        }

        public void SetProperties()
        {
            throw new NotImplementedException();
        }

        public void SaveItem()
        {
            throw new NotImplementedException();
        }
    }
}
