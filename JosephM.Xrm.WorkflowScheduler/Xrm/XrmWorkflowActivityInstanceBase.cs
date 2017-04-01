using System;
using System.Activities;
using JosephM.Xrm.WorkflowScheduler.Core;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;

namespace JosephM.Xrm.WorkflowScheduler
{
    public abstract class XrmWorkflowActivityInstanceBase
    {
        protected XrmWorkflowActivityRegistration Activity { get; set; }

        protected CodeActivityContext ExecutionContext { get; set; }

        private IWorkflowContext _context;

        private IWorkflowContext Context
        {
            get
            {
                if (_context == null)
                {
                    _context = ExecutionContext.GetExtension<IWorkflowContext>();
                    if (_context == null)
                        throw new InvalidPluginExecutionException("Failed to retrieve workflow context.");
                }
                return _context;
            }
        }

        private XrmService _xrmService;

        public XrmService XrmService
        {
            get
            {
                if (_xrmService == null)
                {
                    var serviceFactory = ExecutionContext.GetExtension<IOrganizationServiceFactory>();
                    _xrmService = new XrmService(serviceFactory.CreateOrganizationService(Context.UserId), LogController);
                }
                return _xrmService;
            }
            set { _xrmService = value; }
        }

        private ITracingService _tracingService;

        private ITracingService TracingService
        {
            get
            {
                if (_tracingService == null)
                {
                    _tracingService = ExecutionContext.GetExtension<ITracingService>();
                    if (_tracingService == null)
                        throw new InvalidPluginExecutionException("Failed to retrieve tracing service.");
                }
                return _tracingService;
            }
        }


        private LogController _logController;

        public LogController LogController
        {
            get
            {
                if (_logController == null)
                {
                    _logController = new LogController(new XrmTraceUserInterface(TracingService));
                }
                return _logController;
            }
            set { _logController = value; }
        }

        internal void ExecuteBase(CodeActivityContext executionContext,
            XrmWorkflowActivityRegistration xrmWorkflowActivityRegistration)
        {
            try
            {
                Activity = xrmWorkflowActivityRegistration;
                ExecutionContext = executionContext;

                TracingService.Trace(
                    "Entered Workflow {0}\nActivity Instance Id: {1}\nWorkflow Instance Id: {2}\nCorrelation Id: {3}\nInitiating User: {4}",
                    GetType().Name,
                    ExecutionContext.ActivityInstanceId,
                    ExecutionContext.WorkflowInstanceId,
                    Context.CorrelationId,
                    Context.InitiatingUserId);
                Execute();
            }
            catch (InvalidPluginExecutionException ex)
            {
                LogController.LogLiteral(ex.XrmDisplayString());
                throw;
            }
            catch (Exception ex)
            {
                LogController.LogLiteral(ex.XrmDisplayString());
                throw new InvalidPluginExecutionException(ex.Message, ex);
            }
        }

        protected abstract void Execute();

        private Guid? _targetId;
        public Guid TargetId
        {
            get
            {
                return _targetId.HasValue ? _targetId.Value : Context.PrimaryEntityId;
            }
            set
            {
                _targetId = value;
            }
        }

        private Guid? _currentUserId;
        public Guid CurrentUserId
        {
            get
            {
                return _currentUserId.HasValue ? _currentUserId.Value : Context.InitiatingUserId;
            }
            set
            {
                _currentUserId = value;
            }
        }

        private string _targetType;
        public string TargetType
        {
            get
            {
                return !string.IsNullOrWhiteSpace(_targetType) ? _targetType : Context.PrimaryEntityName;
            }
            set
            {
                _targetType = value;
            }
        }

        public void Trace(string message)
        {
            TracingService.Trace(message);
        }

        public bool IsSandboxIsolated
        {
            get
            {
                return Context.IsolationMode == 2;
            }
        }

        private int _maxSandboxIsolationExecutionSeconds = 120;
        public int MaxSandboxIsolationExecutionSeconds
        {
            get { return _maxSandboxIsolationExecutionSeconds; }
            set { _maxSandboxIsolationExecutionSeconds = value; }
        }
    }
}