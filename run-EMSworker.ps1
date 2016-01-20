# run-emsworker.ps1

# Listen to a RabbitMQ queue for work (RPC requests), handle them with Exchange Management Shell, send the reply.

import-module PSRabbitMQ

echo "I should have access to rabbitmq now."

Set-RabbitMqConfig -computername hare-1.uvm.edu

$credfilename = "c:\local\scripts\emsuser.crd"

# when you don't have a cred file
#
try {
    $password = get-content $credfilename | convertto-securestring

    $credRabbit = new-object system.management.automation.PsCredential("emsuser", $password)

} catch {
    echo "Could not find credential file $credfilename, asking user instead."
    $credRabbit = get-credential -message "Enter user/pass for access to rabbitmq." -username emsuser
}
        
         
$params = @{
    Credential = $credRabbit
    Ssl = "Tls12"
    }


Add-Type -AssemblyName System.Web

# wait for a message

while($true) {
    try {
        [xml] $val = Wait-RabbitMqMessage @params -Exchange "ems_exchange" -QueueName "rpc_queue" -durable:$false -timeout 15 
    } catch [Microsoft.Powershell.Commands.WriteErrorException] {
        #echo "Probably a timeout, going to ignore."
    }

     if ($val) { 
        #$val | fl *
        echo "Callback queue:", $val.request.callback_queue
        echo "Correlation ID:", $val.request.correlation_id
        echo "Command: ", $val.request.command
        try {
                $result = invoke-expression $val.request.command
                $errorCode = $?
        } catch {
                $e = $_.Exception
                $exception_message = $e.Message
        }
        echo "result is $result"
        $correlation_id = [string]$val.request.correlation_id
        $response = "<response><correlation_id>$correlation_id</correlation_id>"
        $response += "<result>"
        
        $response += [System.Web.HttpUtility]::HtmlEncode($result)
        $response += "</result>"
        if ($e) {
            $response += "<exception>$($exception_message)</exception>"
        }
        $response += "<errorCode>$($errorCode)</errorCode>"
        $response += "</response>"

        $callback_queue = [string]$val.request.callback_queue

        echo "going to respond to queue $callback_queue"

        echo "response is $response"

        $response | send-rabbitmqmessage @params -computername hare-1.uvm.edu -exchange "ems_exchange" -key "$callback_queue" -depth 0

        $val = $null  # we've responded, no need to parse it again in the next loop.
        $e = $null
        $result = $null
        $errorCode = $null

    }

}


