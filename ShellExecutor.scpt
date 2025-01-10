on runShellCommand(command)
    try
        -- Run the shell command and return the output
        set result to do shell script command
        return result
    on error errMsg
        return "Error: " & errMsg
    end try
end runShellCommand
