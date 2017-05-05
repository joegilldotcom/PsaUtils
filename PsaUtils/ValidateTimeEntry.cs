using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Activities;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using Microsoft.Xrm.Sdk.Query;

namespace PsaUtils
{

    public class ValidateTimeEntry : CodeActivity
    {
        [RequiredArgument]
        [Input("Resource")]
        [ReferenceTarget("bookableresource")]
        public InArgument<EntityReference> Resource { get; set; }

        [RequiredArgument]
        [Input("Entry Date")]
        public InArgument<DateTime> EntryDate { get; set; }

        [RequiredArgument]
        [Input("Max Minutes per Day")]
        public InArgument<int> MaxMinutes { get; set; }

        [RequiredArgument]
        [Input("Max Minutes Warning Message")]
        public InArgument<string> MaxMinutesWarning { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracer = executionContext.GetExtension<ITracingService>();
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            EntityReference resource = this.Resource.Get(executionContext);
            DateTime entryDate = this.EntryDate.Get(executionContext);
            int maxMinutes = this.MaxMinutes.Get(executionContext);
            string maxWarning = this.MaxMinutesWarning.Get(executionContext);

            string fetchXml = @"<fetch version='1.0' distinct='false' mapping='logical' aggregate='true'>
            <entity name='msdyn_timeentry' >
                <attribute name='msdyn_duration' aggregate='sum' alias ='durationtotal'/>
                <filter type='and' >
                    <condition attribute='msdyn_bookableresource' operator='eq' value='{0}' />
                    <condition attribute = 'msdyn_date' operator= 'on' value = '{1}' />
                </filter>
            </entity>
            </fetch>";

            string formatXml = string.Format(fetchXml, resource.Id.ToString(), entryDate.ToString("yyyy-MM-dd"));
            tracer.Trace(formatXml);

            EntityCollection eColl = service.RetrieveMultiple(new FetchExpression(formatXml));

            foreach (var c in eColl.Entities)
            {
                int? total = ((int?)((AliasedValue)c["durationtotal"]).Value);
                if (total != null & total > maxMinutes)
                       throw new InvalidPluginExecutionException(maxWarning); 
            }
        }
    }
}
