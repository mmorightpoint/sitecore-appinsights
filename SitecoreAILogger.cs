using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using log4net.Appender;
using log4net.spi;
using Microsoft.ApplicationInsights;
using Microsoft.ApplicationInsights.Channel;
using Microsoft.ApplicationInsights.DataContracts;
using Microsoft.ApplicationInsights.Extensibility.Implementation;

namespace Sandbox.ApplicationInsights
{
    public sealed class ApplicationInsightsAppender : AppenderSkeleton
    {
        private TelemetryClient TelemetryClient { get; set; }

        /// <summary> 
        /// The <see cref="T:Microsoft.ApplicationInsights.Log4NetAppender.ApplicationInsightsAppender" /> requires a layout. 
        /// This Appender converts the LoggingEvent it receives into a text string and requires the layout format string to do so. 
        /// </summary> 
        protected override bool RequiresLayout => true;

        /// <summary> 
        /// Initializes the Appender and perform instrumentationKey validation. 
        /// </summary> 
        public override void ActivateOptions()
        {
            base.ActivateOptions();

            TelemetryClient = new TelemetryClient
            {
                Context =
                {
                    InstrumentationKey = "INSTRUMENTATION KEY"
                }
            };

            TelemetryClient.Context.GetInternalContext().SdkVersion = "Log4Net: " + GetAssemblyVersion();
        }

        /// <summary> 
        /// Append LoggingEvent Application Insights logging framework. 
        /// </summary> 
        /// <param name="loggingEvent">Events to be logged.</param> 
        protected override void Append(LoggingEvent loggingEvent)
        {
            if (!string.IsNullOrEmpty(loggingEvent.GetExceptionStrRep()))
            {
                SendException(loggingEvent);
                return;
            }
            SendTrace(loggingEvent);
        }

        private static string GetAssemblyVersion()
        {
            return
                typeof (ApplicationInsightsAppender).Assembly.GetCustomAttributes(false)
                    .OfType<AssemblyFileVersionAttribute>()
                    .First()
                    .Version;
        }

        private static void AddLoggingEventProperty(string key, string value, IDictionary<string, string> metaData)
        {
            if (value != null)
            {
                metaData.Add(key, value);
            }
        }

        private void SendException(LoggingEvent loggingEvent)
        {
            try
            {
                var exception = GetException(loggingEvent);
                var exceptionTelemetry = new ExceptionTelemetry(exception)
                {
                    SeverityLevel = GetSeverityLevel(loggingEvent.Level)
                };
                BuildCustomProperties(loggingEvent, exceptionTelemetry);
                TelemetryClient.Track(exceptionTelemetry);
            }
            catch (ArgumentNullException ex)
            {
                throw new LogException(ex.Message, ex);
            }
        }

        private static Exception GetException(LoggingEvent loggingEvent)
        {
            var exception = GetInstanceField(typeof(LoggingEvent), loggingEvent, "m_thrownException");
            return exception as Exception;
        }

        private static object GetInstanceField(Type type, object instance, string fieldName)
        {
            var field = type.GetField(fieldName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static);
            return field?.GetValue(instance);
        }

        private void SendTrace(LoggingEvent loggingEvent)
        {
            try
            {
                var message = (loggingEvent.RenderedMessage != null) ? RenderLoggingEvent(loggingEvent) : "Log4Net Trace";
                var traceTelemetry = new TraceTelemetry(message)
                {
                    SeverityLevel = GetSeverityLevel(loggingEvent.Level)
                };
                BuildCustomProperties(loggingEvent, traceTelemetry);
                TelemetryClient.Track(traceTelemetry);
            }
            catch (ArgumentNullException ex)
            {
                throw new LogException(ex.Message, ex);
            }
        }

        private static void BuildCustomProperties(LoggingEvent loggingEvent, ITelemetry trace)
        {
            trace.Timestamp = loggingEvent.TimeStamp;
            trace.Context.User.Id = loggingEvent.UserName;

            var telemetry = trace as ExceptionTelemetry;
            var properties = telemetry != null 
                ? telemetry.Properties 
                : ((TraceTelemetry) trace).Properties;

            AddLoggingEventProperty("LoggerName: ", loggingEvent.LoggerName, properties);
            AddLoggingEventProperty("ThreadName: ", loggingEvent.ThreadName, properties);
            

            var locationInformation = loggingEvent.LocationInformation;
            if (locationInformation != null)
            {
                AddLoggingEventProperty("ClassName: ", locationInformation.ClassName, properties);
                AddLoggingEventProperty("FileName: ", locationInformation.FileName, properties);
                AddLoggingEventProperty("MethodName: ", locationInformation.MethodName, properties);
                AddLoggingEventProperty("LineNumber: ", locationInformation.LineNumber, properties);
            }

            AddLoggingEventProperty("Domain: ", loggingEvent.Domain, properties);
            AddLoggingEventProperty("Identity: ", loggingEvent.Identity, properties);

            var eventProperties = loggingEvent.Properties;
            if (eventProperties == null)
                return;

            var keys = eventProperties.GetKeys();
            foreach (var text in keys)
            {
                if (string.IsNullOrEmpty(text) || text.StartsWith("log4net", StringComparison.OrdinalIgnoreCase))
                    continue;

                var obj = eventProperties[text];
                if (obj != null)
                {
                    AddLoggingEventProperty(text, obj.ToString(), properties);
                }
            }
        }

        private static SeverityLevel? GetSeverityLevel(Level logginEventLevel)
        {
            if (logginEventLevel == null)
                return null;

            if (logginEventLevel < Level.INFO)
                return SeverityLevel.Verbose;

            if (logginEventLevel < Level.WARN)
                return SeverityLevel.Information;

            if (logginEventLevel < Level.ERROR)
                return SeverityLevel.Warning;

            if (logginEventLevel < Level.SEVERE)
                return SeverityLevel.Error;

            return SeverityLevel.Critical;
        }
    }
}