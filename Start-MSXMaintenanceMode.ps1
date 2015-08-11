[cmdletBinding()]
param (
    [parameter(Mandatory=$True)][string]$ServerName
)
# From: http://blogs.technet.com/b/nawar/archive/2014/03/30/exchange-2013-maintenance-mode.aspx
Set-PSDebug -Strict

# Procedure for putting Mailbox Servers that are Database Availability Group Members into Maintenance mode? 

# 1. Drain active mail queues on the mailbox server
Set-ServerComponentState $ServerName -Component HubTransport -State Draining -Requester Maintenance

# 2. To help transport services immediately pick the state change run:
# For Mailbox Server role:
Restart-Service MSExchangeTransport 

# If the server is a multi-role server(CAS/MBX) you need to run
Restart-Service MSExchangeTransport 
Restart-Service MSExchangeFrontEndTransport

# 3. To redirect messages pending delivery in the local queues to another Mailbox server run:
Redirect-Message -Server $ServerName -Target <MailboxServerFQDN>
# Note: The target Server value has to be the target server’s FQDN and that the target server shouldn’t be in maintenance mode.

# 4. To prevents the node from being and becoming the PAM, pause the cluster node by running
Suspend-ClusterNode $ServerName

# 5. To move all active databases currently hosted on the DAG member to other DAG members, run
Set-MailboxServer $ServerName -DatabaseCopyActivationDisabledAndMoveNow $True

# 6. Get the status of the existing database copy auto activation policy, run the following and note the value of DatabaseCopyAutoActivationPolicy, we will need this when taking the server out of Maintenance in the future
Get-MailboxServer $ServerName | Select DatabaseCopyAutoActivationPolicy
# To prevent the server from hosting active database copies, run
Set-MailboxServer $ServerName -DatabaseCopyAutoActivationPolicy Blocked

#7. To put the server in maintenance mode run:
Set-ServerComponentState $ServerName -Component ServerWideOffline -State Inactive -Requester Maintenance

# Note: Closely monitor the transport queue before running the step above , queues at this stage should be empty or nearly empty, as we will be disabling all server components, any mails still pending in the queues will have delay in delivery till the server is taken out from maintenance mode.