/*================================================================================================================================

  This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment.  

  THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, 
  INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  

  We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute the object 
  code form of the Sample Code, provided that You agree: (i) to not use Our name, logo, or trademarks to market Your software 
  product in which the Sample Code is embedded; (ii) to include a valid copyright notice on Your software product in which the 
  Sample Code is embedded; and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any claims 
  or lawsuits, including attorneys’ fees, that arise or result from the use or distribution of the Sample Code.

 =================================================================================================================================*/

using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;
using Microsoft.Xrm.Tooling.Connector;

namespace Microsoft.Pfe.Xrm.Helper
{
    /// <summary>
    /// Extended OrganiztaionService which exposed PreExecute event to obtain 
    /// QueryExpression before executing RetrieveMutliple
    /// </summary>
    public class OrganizationServiceEx : IOrganizationService
    {
        private CrmServiceClient _service;
        public event EventHandler<PreExecuteEventArgs> PreExecute;

        /// <summary>
        /// Constructor. Pass CrmConnection to base.
        /// </summary>
        /// <param name="cnn">CrmConnection</param>
        public OrganizationServiceEx(CrmServiceClient cnn)
        {
            this._service = cnn;
        }

        public Guid Create(Entity entity)
        {
            return _service.Create(entity);
        }

        public Entity Retrieve(string entityName, Guid id, ColumnSet columnSet)
        {
            return _service.Retrieve(entityName,id,columnSet);
        }

        public void Update(Entity entity)
        {
            _service.Update(entity);
        }

        public void Delete(string entityName, Guid id)
        {
            _service.Delete(entityName,id);
        }

        /// <summary>
        /// Override Execute method only
        /// </summary>
        /// <param name="request">OrganizationRequest</param>
        /// <returns>OrganizationResponse</returns>
        public OrganizationResponse Execute(OrganizationRequest request)
        {
            // If request is RetrieveMultipleReuqest, and Query is QueryExpression, then propagate to Event.
            if (request is RetrieveMultipleRequest && (request as RetrieveMultipleRequest).Query is QueryExpression)
            {
                this.PreExecute(this, new PreExecuteEventArgs((request as RetrieveMultipleRequest).Query as QueryExpression));
            }
            return _service.Execute(request);
        }

        public void Associate(string entityName, Guid entityId, Relationship relationship, EntityReferenceCollection relatedEntities)
        {
            _service.Associate(entityName, entityId, relationship, relatedEntities);
        }

        public void Disassociate(string entityName, Guid entityId, Relationship relationship, EntityReferenceCollection relatedEntities)
        {
            _service.Disassociate(entityName, entityId, relationship, relatedEntities);
        }

        public EntityCollection RetrieveMultiple(QueryBase query)
        {
            return _service.RetrieveMultiple(query);
        }
    }

    /// <summary>
    /// PreExecute EventArgs which has QueryExpression.
    /// </summary>
    public class PreExecuteEventArgs : EventArgs
    {
        public QueryExpression query { get; set; }
        public PreExecuteEventArgs(QueryExpression query)
        {
            this.query = query;
        }
    }
}
