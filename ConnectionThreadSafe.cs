using Microsoft.Xrm.Sdk.Client;
using System;
using System.Configuration;

namespace PFC.Tools.CRMConnection
{
    public static class ConnectionThreadSafe
    {
        private static ServerConnection.Configuration config;

        public static OrganizationServiceProxy GetOrganizationProxy(string user, string password, string serverAddress, string ssl, string organizationName, string isO365Org)
        {
            try
            {
                if (config == null)
                {
                    ServerConnection serverConnect = new ServerConnection();
                    config = serverConnect.GetServerConfiguration(user, password, serverAddress, ssl, organizationName, isO365Org);
                }
                else if (config.OrganizationName != organizationName || config.Credentials.UserName.UserName != user)
                {                  
                    ServerConnection serverConnect = new ServerConnection();
                    config = serverConnect.GetServerConfiguration(user, password, serverAddress, ssl, organizationName, isO365Org);
                }
                else
                {

                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return ServerConnection.GetOrganizationProxy(config);
        }

        public static OrganizationServiceProxy GetOrganizationProxy(ConnectionData connection)
        {
            return GetOrganizationProxy(
                connection.User,
                connection.Password,
                connection.ServerAddress,
                connection.SSL ? "Y" : "N",
                connection.OrganizationName,
                connection.IsO365Org ? "Y" : "N"
                );
        }
    }

    public static class PrayerCrmServerExtension
    {
        public static Guid PrayerCreate(this OrganizationServiceProxy service, Microsoft.Xrm.Sdk.Entity entity)
        {
            TimeSpan timeout = TimeSpan.FromSeconds(10);
            return service.PrayerCreate(entity, timeout, 5);
        }
        public static Guid PrayerCreate(this OrganizationServiceProxy service, Microsoft.Xrm.Sdk.Entity entity, TimeSpan timeout, int tryCount)
        {
            TimeSpan tempTimeOut = service.Timeout;
            service.Timeout = timeout;

            int tc = tryCount;

            do
            {
                try
                {
                    if (tc == 1) service.Timeout = tempTimeOut;
                    var id = service.Create(entity);

                    service.Timeout = tempTimeOut;
                    return id;
                }
                catch (Exception e)
                {
                    if (e.Message == "Generic SQL error." ||
                        e is System.Data.SqlClient.SqlException ||
                        e is TimeoutException)
                    {
                        tc--;
                        if (tryCount <= 0)
                        {
                            service.Timeout = tempTimeOut;
                            throw e;
                        }
                    }
                    else
                    {
                        service.Timeout = tempTimeOut;
                        throw e;
                    }
                }

            } while (tryCount > 0);

            return Guid.Empty;
        }

        public static void PrayerUpdate(this OrganizationServiceProxy service, Microsoft.Xrm.Sdk.Entity entity)
        {
            TimeSpan timeout = TimeSpan.FromSeconds(10);
            service.PrayerUpdate(entity, timeout, 5);
        }
        public static void PrayerUpdate(this OrganizationServiceProxy service, Microsoft.Xrm.Sdk.Entity entity, TimeSpan timeout, int tryCount)
        {
            TimeSpan tempTimeOut = service.Timeout;
            service.Timeout = timeout;

            int tc = tryCount;

            do
            {
                try
                {
                    if (tc == 1) service.Timeout = tempTimeOut;
                    service.Update(entity);

                    service.Timeout = tempTimeOut;
                    return;
                }
                catch (Exception e)
                {
                    if (e.Message == "Generic SQL error." ||
                        e is System.Data.SqlClient.SqlException ||
                        e is TimeoutException)
                    {
                        tc--;
                        if (tryCount <= 0)
                        {
                            service.Timeout = tempTimeOut;
                            throw e;
                        }
                    }
                    else
                    {
                        service.Timeout = tempTimeOut;
                        throw e;
                    }
                }

            } while (tryCount > 0);

            return;
        }

        public static Microsoft.Xrm.Sdk.Entity PrayerRetrieve(this OrganizationServiceProxy service, string entityName, Guid id, Microsoft.Xrm.Sdk.Query.ColumnSet columnSet)
        {
            TimeSpan timeout = TimeSpan.FromSeconds(10);
            return service.PrayerRetrieve(entityName, id, columnSet, timeout, 5);
        }
        public static Microsoft.Xrm.Sdk.Entity PrayerRetrieve(this OrganizationServiceProxy service, string entityName, Guid id, Microsoft.Xrm.Sdk.Query.ColumnSet columnSet, TimeSpan timeout, int tryCount)
        {
            TimeSpan tempTimeOut = service.Timeout;
            service.Timeout = timeout;

            int tc = tryCount;

            do
            {
                try
                {
                    if (tc == 1) service.Timeout = tempTimeOut;
                    var entity = service.Retrieve(entityName, id, columnSet);

                    service.Timeout = tempTimeOut;
                    return entity;
                }
                catch (Exception e)
                {
                    if (e.Message == "Generic SQL error." ||
                        e is System.Data.SqlClient.SqlException ||
                        e is TimeoutException)
                    {
                        tc--;
                        if (tryCount <= 0)
                        {
                            service.Timeout = tempTimeOut;
                            throw e;
                        }
                    }
                    else
                    {
                        service.Timeout = tempTimeOut;
                        throw e;
                    }
                }

            } while (tryCount > 0);

            return null;
        }

        public static Microsoft.Xrm.Sdk.EntityCollection PrayerRetrieveMultiple(this OrganizationServiceProxy service, Microsoft.Xrm.Sdk.Query.QueryBase query)
        {
            TimeSpan timeout = TimeSpan.FromSeconds(25);
            return service.PrayerRetrieveMultiple(query, timeout, 2);
        }
        public static Microsoft.Xrm.Sdk.EntityCollection PrayerRetrieveMultiple(this OrganizationServiceProxy service, Microsoft.Xrm.Sdk.Query.QueryBase query, TimeSpan timeout, int tryCount)
        {
            TimeSpan tempTimeOut = service.Timeout;
            service.Timeout = timeout;

            int tc = tryCount;

            do
            {
                try
                {
                    if (tc == 1) service.Timeout = tempTimeOut;
                    var entityCollection = service.RetrieveMultiple(query);

                    service.Timeout = tempTimeOut;
                    return entityCollection;
                }
                catch (Exception e)
                {
                    if (e.Message == "Generic SQL error." ||
                        e is System.Data.SqlClient.SqlException ||
                        e is TimeoutException)
                    {
                        tc--;
                        if (tryCount <= 0)
                        {
                            service.Timeout = tempTimeOut;
                            throw e;
                        }
                    }
                    else
                    {
                        service.Timeout = tempTimeOut;
                        throw e;
                    }
                }

            } while (tryCount > 0);

            return null;
        }
    }

    public class ConnectionData
    {
        public string ServerAddress { get; set; }
        public string OrganizationName { get; set; }
        public bool SSL { get; set; }
        public bool IsO365Org { get; set; }
        public string User { get; set; }
        public string Password { get; set; }
    }
}
