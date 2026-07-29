const AGENDA_DAYS = [
  {
    id: 'day-one',
    label: 'Day One',
    date: 'Tuesday, September 1, 2026',
    theme: 'Arrival + Welcome',
    sessions: [
      {
        start: '4:00 PM',
        end: '7:00 PM',
        track: 'Registration Desk',
        room: 'Registration',
        type: 'Registration',
        title: 'Registration Open - All Attendees'
      },
      {
        start: '5:00 PM',
        end: '7:00 PM',
        track: 'Expo Hall',
        room: 'Networking',
        type: 'Networking',
        title: 'Welcome Reception & Expo Hall'
      }
    ]
  },
  {
    id: 'day-two',
    label: 'Day Two',
    date: 'Wednesday, September 2, 2026',
    theme: 'Market Shifts, AI + Operating Smarter',
    sessions: [
      { start: '8:00 AM', end: '9:00 AM', track: 'General Session', room: 'GB 3&4', type: 'Meal / Expo', title: 'Breakfast & Expo' },
      { start: '9:00 AM', end: '9:15 AM', track: 'General Session', room: 'GB 3&4', type: 'General Session', title: 'Welcome to Rev.io Summit' },
      { start: '9:15 AM', end: '10:00 AM', track: 'General Session', room: 'GB 3&4', type: 'General Session', title: 'The Next Era of Service Providers: Operating Smarter with Rev.io' },
      { start: '10:00 AM', end: '10:05 AM', track: 'General Session', room: 'GB 3&4', type: 'Sponsor Session', title: 'Diamond Sponsor Keynote' },
      { start: '10:05 AM', end: '10:30 AM', track: 'General Session', room: 'GB 3&4', type: 'Break', title: 'Coffee Break & Expo' },
      { start: '10:30 AM', end: '11:00 AM', track: 'General Session', room: 'GB 3&4', type: 'Panel', title: 'How the Service Provider Market Is Shifting - and What Operators Should Do Next' },
      { start: '11:00 AM', end: '11:45 AM', track: 'General Session', room: 'GB 3&4', type: 'Keynote', title: 'Jay McBain Keynote' },
      { start: '11:45 AM', end: '1:00 PM', track: 'Garden Court', room: 'GB 3&4', type: 'Meal / Expo', title: 'Lunch & Expo: GreatAmerica Keynote' },
      { start: '1:00 PM', end: '1:30 PM', track: 'Track A', room: 'GB 1', type: 'Sponsor Session', title: 'Sponsor Session' },
      { start: '1:00 PM', end: '1:30 PM', track: 'Track B', room: 'GB 2', type: 'Sponsor Session', title: 'Sponsor Panel' },
      { start: '1:30 PM', end: '1:35 PM', track: 'Garden Court', room: 'Garden Court', type: 'Break', title: '5 Minute Break' },
      { start: '1:35 PM', end: '2:20 PM', track: 'Track A', room: 'GB 1', type: 'Breakout', title: 'Billing Best Practices: Turning Complex Service Revenue into Clean, Repeatable Operations' },
      { start: '1:35 PM', end: '2:20 PM', track: 'Track B', room: 'GB 2', type: 'Breakout', title: 'AI-Ready and Security-Ready: Preparing Your Service Provider Business for the Next Operating Model (Cynomi)' },
      { start: '2:20 PM', end: '2:30 PM', track: 'Garden Court', room: 'Garden Court', type: 'Break', title: '10 Minute Break' },
      { start: '2:30 PM', end: '3:15 PM', track: 'Track A', room: 'GB 1', type: 'Breakout', title: 'From RMM Alert to Resolved Ticket: Automating the Service Desk Workflow (NinjaOne)' },
      { start: '2:30 PM', end: '3:15 PM', track: 'Track B', room: 'GB 2', type: 'Breakout', title: 'Meeting Rising Customer Expectations with a Modern Service Desk' },
      { start: '2:30 PM', end: '3:15 PM', track: 'Track C', room: 'Wilton', type: 'Breakout', title: 'From Quote to Revenue: Growing Your Service Portfolio with CommerceHub' },
      { start: '3:15 PM', end: '3:30 PM', track: 'Garden Court', room: 'Garden Court', type: 'Break', title: 'Break & Expo' },
      { start: '3:30 PM', end: '4:00 PM', track: 'Track A', room: 'GB 1', type: 'Breakout', title: 'The Service Provider KPI Playbook: Proving Operational Health and Business Value' },
      { start: '3:30 PM', end: '4:00 PM', track: 'Track B', room: 'GB 2', type: 'Breakout', title: 'Managing Complex Customer Projects Without Losing Control' },
      { start: '3:30 PM', end: '4:00 PM', track: 'Track C', room: 'Wilton', type: 'Breakout', title: 'Your QBR is Bullsh*t - Delivering Tangible Value to Your Customers' },
      { start: '4:00 PM', end: '4:10 PM', track: 'Garden Court', room: 'Garden Court', type: 'Break', title: '10 Minute Break & Expo' },
      { start: '4:10 PM', end: '4:50 PM', track: 'General Session', room: 'GB 3&4', type: 'General Session', title: 'Revii Showcase: Your AI-Powered Digital Workforce in Action' },
      { start: '4:50 PM', end: '5:00 PM', track: 'General Session', room: 'GB 3&4', type: 'General Session', title: 'Day Two Closing Remarks' },
      { start: '5:00 PM', end: '6:30 PM', track: 'Break', room: 'Garden Court', type: 'Break', title: 'All-Attendee Break (no expo)' },
      { start: '6:30 PM', end: '10:00 PM', track: 'Networking Reception', room: 'Truist Park (offsite)', type: 'Networking', title: 'Networking Reception - Braves Stadium (sponsored by CCH SureTax)' }
    ]
  },
  {
    id: 'day-three',
    label: 'Day Three',
    date: 'Thursday, September 3, 2026',
    theme: 'Roadmap, Workshops + Closing',
    sessions: [
      { start: '8:00 AM', end: '9:00 AM', track: 'General Session', room: 'GB 3&4', type: 'Meal / Expo', title: 'Breakfast - Sponsor Keynote Spotlight: Liongard' },
      { start: '9:00 AM', end: '9:30 AM', track: 'General Session', room: 'GB 3&4', type: 'General Session', title: 'Opening Remarks' },
      { start: '9:30 AM', end: '10:15 AM', track: 'General Session', room: 'GB 3&4', type: 'General Session', title: "Rev.io Product Roadmap: What's Next for the Service Provider Operating Platform" },
      { start: '10:15 AM', end: '10:30 AM', track: 'General Session', room: 'GB 3&4', type: 'Break', title: 'Break & Expo' },
      { start: '10:15 AM', end: '10:45 AM', track: 'General Session', room: 'GB 3&4', type: 'General Session', title: 'Billing for AI Monetization: Packaging and Charging for the New Digital Workforce' },
      { start: '10:45 AM', end: '11:15 AM', track: 'General Session', room: 'GB 3&4', type: 'Workshop', title: 'Onboarding Readiness Workshop: Set Up Your Team, Data, and Workflows for Success' },
      { start: '11:15 AM', end: '11:30 AM', track: 'General Session', room: 'GB 3&4', type: 'Break', title: '15 Minute Break & Expo' },
      { start: '11:30 AM', end: '12:00 PM', track: 'Track A', room: 'GB 1', type: 'Breakout', title: 'Rev.io Rookies Roundtable' },
      { start: '11:30 AM', end: '12:00 PM', track: 'Track B', room: 'GB 2', type: 'Breakout', title: 'Field Service Without the Desk: Running Technician Workflows from the Rev.io Mobile App' },
      { start: '12:00 PM', end: '1:05 PM', track: 'General Session', room: 'GB 3&4', type: 'Meal / Expo', title: 'Lunch & Expo' },
      { start: '1:05 PM', end: '1:35 PM', track: 'Track A', room: 'GB 1', type: 'Sponsor Session', title: 'Sponsor Panel' },
      { start: '1:05 PM', end: '1:35 PM', track: 'Track B', room: 'GB 2', type: 'Sponsor Session', title: 'Sponsor Panel' },
      { start: '1:35 PM', end: '1:45 PM', track: 'Garden Court', room: 'GB 3&4', type: 'Break', title: '10 Minute Break & Expo' },
      { start: '1:45 PM', end: '2:15 PM', track: 'Track A', room: 'GB 1', type: 'Breakout', title: 'Sales Best Practices for Service Providers: Turning Better Process into Better Revenue' },
      { start: '1:45 PM', end: '2:15 PM', track: 'Track B', room: 'GB 2', type: 'Workshop', title: 'Forge Showcase - Live Workshop' },
      { start: '2:15 PM', end: '2:25 PM', track: 'Garden Court', room: 'GB 3&4', type: 'Break', title: '10 Minute Break & Expo' },
      { start: '2:25 PM', end: '2:55 PM', track: 'Track A', room: 'GB 1', type: 'Breakout', title: 'Business Insights in Rev.io: Turning Operational Data into Better Decisions' },
      { start: '2:25 PM', end: '2:55 PM', track: 'Track B', room: 'GB 2', type: 'Breakout', title: 'What Is a Modern Service Provider? The Operating Model for the Next Era' },
      { start: '2:55 PM', end: '3:00 PM', track: 'Garden Court', room: 'GB 3&4', type: 'Break', title: '5 Minute Break' },
      { start: '3:00 PM', end: '3:15 PM', track: 'General Session', room: 'GB 3&4', type: 'General Session', title: 'Closing Remarks + Save the Date for Summit 2027' }
    ]
  }
];

module.exports = { AGENDA_DAYS };
