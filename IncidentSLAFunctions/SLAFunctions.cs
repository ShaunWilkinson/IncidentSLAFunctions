using System;
using System.Activities;

using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;

namespace IncidentSLAFunctions
{
    
    // test
    // Calculates the elapsed time between the start date and the current date within a case
    public class GetElapsedTime : CodeActivity
    {
        #region inputs & outputs
        [RequiredArgument]
        [Input("Case")]
        [ReferenceTarget("incident")]
        public InArgument<EntityReference> EnteredCase { get; set; }

        [Output("Actual Seconds")]
        public OutArgument<int> ElapsedTime { get; set; }
        #endregion inputs & outputs

        protected override void Execute(CodeActivityContext executionContext)
        {
            #region Initiate needed context and service

            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory =
                executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service =
                serviceFactory.CreateOrganizationService(context.UserId);

            #endregion Initiate needed context and service

            #region Retrieve Case Details
            Guid caseId = EnteredCase.Get<EntityReference>(executionContext).Id;

            // Retrieve the related case
            RetrieveRequest request = new RetrieveRequest();
            request.ColumnSet = new ColumnSet("new_slastartdate");
            request.Target = new EntityReference("incident", caseId);

            Entity caseEntity = (Entity)((RetrieveResponse)service.Execute(request)).Entity;
            #endregion Retrieve Case Details

            #region Calculate elapsed time and return
            DateTime now = DateTime.Now;
            DateTime startTime = caseEntity.GetAttributeValue<DateTime>("new_slastartdate");

            TimeSpan calculatedTime = now - startTime;
            Int32 elapsedTime = Convert.ToInt32(calculatedTime.TotalSeconds);

            ElapsedTime.Set(executionContext, elapsedTime);
            #endregion Calculate elapsed time and return
        }
    }

    // Calculates a new date time by subtracting seconds from a given DateTime
    public class SubtractSecondsFromDateTime : CodeActivity
    {
        #region inputs & outputs
        [RequiredArgument]
        [Input("DateTime")]
        public InArgument<DateTime> DateTimeVal { get; set; }

        [RequiredArgument]
        [Input("Seconds to Subtract")]
        public InArgument<int> Seconds { get; set; }

        [Output("Calculated DateTime")]
        public OutArgument<DateTime> CalculatedValue { get; set; }
        #endregion inputs & outputs

        protected override void Execute(CodeActivityContext context)
        {
            DateTime dateTimeInput = DateTimeVal.Get<DateTime>(context);
            int seconds = Seconds.Get<int>(context);

            DateTime calculatedDate = dateTimeInput.AddSeconds(seconds * -1);
            
            CalculatedValue.Set(context, calculatedDate);
        }
    }
}
