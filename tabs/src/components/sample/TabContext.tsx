import {  useState, useEffect } from "react";
import "./Welcome.css";
import { app } from "@microsoft/teams-js";

export const TabContext = () => {

    const [channelName, setChannelName] = useState<string>("");
    const [teamName, setTeamName] = useState<string>("");
    const [host, setHost] = useState<string>("");

    useEffect(() => {
        app.getContext().then( (context: app.Context) => {
            const channelName = context.channel ? context.channel?.displayName : "";
            setChannelName(channelName!);
            const teamName = context.team ? context.team?.displayName: "";
            setTeamName(teamName!);
            setHost(context.app.host.name);
        });

    }, []);
    
    if (host === 'Teams') {    
            return (
                <div>
                    <p className="center"><strong>Context: {host}</strong></p>
                    <p className="center"><strong>Team name: </strong>{teamName}</p>
                    <p className="center"><strong>Channel name: </strong>{channelName}</p>
                </div>
        );
    }
    
    return (<p>
                <p className="center"><strong>Context: {host}</strong></p>
            </p>);
  }