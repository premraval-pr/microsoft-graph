import React from 'react';
import { NavLink as RouterNavLink } from 'react-router-dom';
import { Table } from 'reactstrap';
import moment, { Moment } from 'moment-timezone';
import { findIana } from "windows-iana";
import { Event } from 'microsoft-graph';
import { config } from './Config';
import { getUserWeekCalendar } from './GraphService';
import withAuthProvider, { AuthComponentProps } from './AuthProvider';

interface CalendarState {
  eventsLoaded: boolean;
  events: Event[];
  startOfWeek: Moment | undefined;
}

class Calendar extends React.Component<AuthComponentProps, CalendarState> {
  constructor(props: any) {
    super(props);

    this.state = {
      eventsLoaded: false,
      events: [],
      startOfWeek: undefined
    };
  }

  async componentDidUpdate() {
    if (this.props.user && !this.state.eventsLoaded)
    {
      try {
        // Get the user's access token
        var accessToken = await this.props.getAccessToken(config.scopes);

        // Convert user's Windows time zone ("Pacific Standard Time")
        // to IANA format ("America/Los_Angeles")
        // Moment needs IANA format
        var ianaTimeZone = findIana(this.props.user.timeZone);

        // Get midnight on the start of the current week in the user's timezone,
        // but in UTC. For example, for Pacific Standard Time, the time value would be
        // 07:00:00Z
        var startOfWeek = moment.tz(ianaTimeZone![0].valueOf()).startOf('week').utc();

        // Get the user's events
        var events = await getUserWeekCalendar(accessToken, this.props.user.timeZone, startOfWeek);

        // Update the array of events in state
        this.setState({
          eventsLoaded: true,
          events: events,
          startOfWeek: startOfWeek
        });
      }
      catch (err) {
        this.props.setError('ERROR', JSON.stringify(err));
      }
    }
  }

  render() {
    return (
      <pre><code>{JSON.stringify(this.state.events, null, 2)}</code></pre>
    );
  }
}

export default withAuthProvider(Calendar);