const AGENDA_DAYS = [
  {
    id: 'day-two',
    label: 'Day Two',
    date: 'Wednesday, September 2',
    theme: 'Agenda',
    sessions: [
      { start: '9:00 AM', end: '9:15 AM', title: 'Welcome to Rev.io Summit' },
      { start: '9:15 AM', end: '10:00 AM', title: 'The Next Era of Service Providers: Operating Smarter with Rev.io' },
      { start: '10:00 AM', end: '10:05 AM', title: 'SPONSOR SESSION: Pax8 - Diamond Sponsor Keynote (5 min)' },
      { start: '10:05 AM', end: '10:30 AM', title: 'Coffee Break & Expo' },
      { start: '10:30 AM', end: '11:00 AM', title: 'How the Service Provider Market Is Shifting - and What Operators Should Do Next (Pax8, GTIA, ELT)' },
      { start: '11:00 AM', end: '11:45 AM', title: 'Jay McBain Keynote' },
      { start: '11:45 AM', end: '1:00 PM', title: 'LUNCH & EXPO: GreatAmerica Keynote' },
      { start: '1:00 PM', end: '1:30 PM', title: 'SPONSOR SESSION: CCH Suretax' },
      { start: '1:00 PM', end: '1:30 PM', title: 'SPONSOR PANEL: Sangoma' },
      { start: '1:30 PM', end: '1:35 PM', title: '5 min Break' },
      { start: '1:35 PM', end: '2:20 PM', title: 'Billing Best Practices: Turning Complex Service Revenue into Clean, Repeatable Operations' },
      { start: '1:35 PM', end: '2:20 PM', title: 'AI-Ready and Security-Ready: Preparing Your Service Provider Business for the Next Operating Model (Cynomi)' },
      { start: '2:20 PM', end: '2:30 PM', title: '10 min Break' },
      { start: '2:30 PM', end: '3:15 PM', title: 'From RMM Alert to Resolved Ticket: Automating the Service Desk Workflow (Ninja-One)' },
      { start: '2:30 PM', end: '3:15 PM', title: 'Meeting Rising Customer Expectations with a Modern Service Desk' },
      { start: '2:30 PM', end: '3:15 PM', title: 'CommerceHub in Practice: Building a Smarter Marketplace for Service Provider Growth' },
      { start: '3:15 PM', end: '3:30 PM', title: 'Break & Expo' },
      { start: '3:30 PM', end: '4:00 PM', title: 'The Service Provider KPI Playbook: Proving Operational Health and Business Value' },
      { start: '3:30 PM', end: '4:00 PM', title: 'Managing Complex Customer Projects Without Losing Control' },
      { start: '3:30 PM', end: '4:00 PM', title: 'Your QBR is Bullsh*t - Delivering tangible value to your customers.' },
      { start: '4:00 PM', end: '4:10 PM', title: '10 min Break & Expo' },
      { start: '4:10 PM', end: '4:50 PM', title: 'Revii Showcase: Your AI-Powered Digital Workforce in Action' },
      { start: '4:50 PM', end: '5:00 PM', title: 'REV.IO (Closing) - Day 1 - Closing Remarks' },
      { start: '5:00 PM', end: '6:30 PM', title: 'All Attendee Break; no expo' },
      { start: '6:30 PM', end: '10:00 PM', title: 'Networking Reception - Braves Stadium - CCH Suretax sponsored' }
    ]
  },
  {
    id: 'day-three',
    label: 'Day Three',
    date: 'Thursday, September 3',
    theme: 'Agenda',
    sessions: [
      { start: '8:00 AM', end: '9:00 AM', title: 'Breakfast - SPONSOR Keynote Spotlight: Liongard' },
      { start: '9:00 AM', end: '9:30 AM', title: '(Opening) - Day 2 Opening Remarks + Awards' },
      { start: '9:30 AM', end: '10:15 AM', title: "Rev.io Product Roadmap: What's Next for the Service Provider Operating Platform" },
      { start: '10:15 AM', end: '10:30 AM', title: 'Break & Expo' },
      { start: '10:15 AM', end: '10:45 AM', title: 'Billing for AI Monetization: Packaging and Charging for the New Digital Workforce' },
      { start: '10:45 AM', end: '11:15 AM', title: 'Onboarding Readiness Workshop: Set Up Your Team, Data, and Workflows for Success' },
      { start: '11:15 AM', end: '11:30 AM', title: '15 min Break & Expo' },
      { start: '11:30 AM', end: '12:00 PM', title: 'Rev.io Rookies Roundtable' },
      { start: '11:30 AM', end: '12:00 PM', title: 'Field Service Without the Desk: Running Technician Workflows from the Rev.io Mobile App' },
      { start: '12:00 PM', end: '1:05 PM', title: 'Lunch & Expo' },
      { start: '1:05 PM', end: '1:35 PM', title: 'SPONSOR PANEL: TaxConnex' },
      { start: '1:05 PM', end: '1:35 PM', title: 'SPONSOR PANEL: TaxTheta' },
      { start: '1:35 PM', end: '1:45 PM', title: '10 min Break & Expo' },
      { start: '1:45 PM', end: '2:15 PM', title: 'Sales Best Practices for Service Providers: Turning Better Process into Better Revenue' },
      { start: '1:45 PM', end: '2:15 PM', title: 'Forge Showcase - Live Workshop' },
      { start: '2:15 PM', end: '2:25 PM', title: '10 min Break & Expo' },
      { start: '2:25 PM', end: '2:55 PM', title: 'Business Insights in Rev.io: Turning Operational Data into Better Decisions' },
      { start: '2:25 PM', end: '2:55 PM', title: 'What Is a Modern Service Provider? The Operating Model for the Next Era' },
      { start: '2:55 PM', end: '3:00 PM', title: '5 min Break' },
      { start: '3:00 PM', end: '3:15 PM', title: 'REV.IO (Closing ) - Closing Remarks, Thank you, & sign up for 2027' }
    ]
  }
];

module.exports = { AGENDA_DAYS };
