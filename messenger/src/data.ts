import type { CallRecord, Contact, Conversation } from "./types";

export const ME = {
  id: "me",
  name: "Muhammad",
  handle: "@muhammad",
};

const minutesAgo = (minutes: number) => Date.now() - minutes * 60_000;

export const seedConversations: Conversation[] = [
  {
    id: "aisha",
    title: "Aisha Khan",
    subtitle: "Design review",
    color: "#2f6fed",
    participants: ["aisha"],
    unread: 2,
    messages: [
      {
        id: "a1",
        senderId: "aisha",
        text: "Can you glance at the listing photos before I upload?",
        sentAt: minutesAgo(95),
      },
      {
        id: "a2",
        senderId: "me",
        text: "Send the square crop and I’ll mark the ones that read at thumbnail size.",
        sentAt: minutesAgo(88),
      },
      {
        id: "a3",
        senderId: "aisha",
        text: "Dropped them in the folder. The neon sheet is photo 01.",
        sentAt: minutesAgo(12),
      },
      {
        id: "a4",
        senderId: "aisha",
        text: "Also — pin copy is in the doc if you want a second pass.",
        sentAt: minutesAgo(8),
      },
    ],
  },
  {
    id: "studio",
    title: "Saturday studio",
    subtitle: "Aisha, Jordan, Maya",
    color: "#c45ad4",
    participants: ["aisha", "jordan", "maya"],
    isGroup: true,
    unread: 1,
    messages: [
      {
        id: "s1",
        senderId: "jordan",
        text: "I can hold the loft from 11. Who’s bringing the lights?",
        sentAt: minutesAgo(240),
      },
      {
        id: "s2",
        senderId: "maya",
        text: "I’ll grab the two softboxes. Muhammad, can you bring the backdrop roll?",
        sentAt: minutesAgo(220),
      },
      {
        id: "s3",
        senderId: "me",
        text: "Yes — cream and forest. I’ll be there at 10:45.",
        sentAt: minutesAgo(200),
      },
      {
        id: "s4",
        senderId: "aisha",
        text: "Coffee order: two cortados, one oat latte.",
        sentAt: minutesAgo(34),
      },
    ],
  },
  {
    id: "jordan",
    title: "Jordan Lee",
    subtitle: "Usually replies fast",
    color: "#1aa37a",
    participants: ["jordan"],
    unread: 0,
    messages: [
      {
        id: "j1",
        senderId: "me",
        text: "Did the printer confirm the sticker sheet size?",
        sentAt: minutesAgo(1500),
      },
      {
        id: "j2",
        senderId: "jordan",
        text: "They can do 10×10 on vinyl. Waterproof, same as last run.",
        sentAt: minutesAgo(1480),
      },
      {
        id: "j3",
        senderId: "me",
        text: "Perfect. Let’s lock 200 sheets.",
        sentAt: minutesAgo(1460),
      },
      {
        id: "j4",
        senderId: "jordan",
        text: "Locked. I’ll send the proof tonight.",
        sentAt: minutesAgo(400),
      },
    ],
  },
  {
    id: "maya",
    title: "Maya Patel",
    subtitle: "Family",
    color: "#e07a2f",
    participants: ["maya"],
    unread: 0,
    messages: [
      {
        id: "m1",
        senderId: "maya",
        text: "Are you still coming by Friday? I made extra daal.",
        sentAt: minutesAgo(2800),
      },
      {
        id: "m2",
        senderId: "me",
        text: "Yes — landing around 7. Don’t wait on me for dinner though.",
        sentAt: minutesAgo(2760),
      },
      {
        id: "m3",
        senderId: "maya",
        text: "I’ll keep a plate. Drive safe.",
        sentAt: minutesAgo(900),
      },
    ],
  },
  {
    id: "alex",
    title: "Alex Rivera",
    subtitle: "Shipping",
    color: "#5b67e8",
    participants: ["alex"],
    unread: 0,
    messages: [
      {
        id: "x1",
        senderId: "alex",
        text: "Label for the cookie cutter sample is in your inbox.",
        sentAt: minutesAgo(4200),
      },
      {
        id: "x2",
        senderId: "me",
        text: "Got it. Dropped it at the counter this morning.",
        sentAt: minutesAgo(3900),
      },
    ],
  },
  {
    id: "sam",
    title: "Sam Chen",
    subtitle: "Quiet",
    color: "#0f9d8e",
    participants: ["sam"],
    muted: true,
    unread: 0,
    messages: [
      {
        id: "c1",
        senderId: "sam",
        text: "Thanks for the invoice template notes. The totals block is much clearer.",
        sentAt: minutesAgo(6200),
      },
      {
        id: "c2",
        senderId: "me",
        text: "Glad it helped. Ping me if the A4 crop still feels tight.",
        sentAt: minutesAgo(6100),
      },
    ],
  },
];

export const seedContacts: Contact[] = [
  { id: "aisha", name: "Aisha Khan", color: "#2f6fed", handle: "@aisha", conversationId: "aisha" },
  { id: "jordan", name: "Jordan Lee", color: "#1aa37a", handle: "@jordan", conversationId: "jordan" },
  { id: "maya", name: "Maya Patel", color: "#e07a2f", handle: "@maya", conversationId: "maya" },
  { id: "alex", name: "Alex Rivera", color: "#5b67e8", handle: "@alex", conversationId: "alex" },
  { id: "sam", name: "Sam Chen", color: "#0f9d8e", handle: "@sam", conversationId: "sam" },
];

export const seedCalls: CallRecord[] = [
  {
    id: "call-1",
    name: "Aisha Khan",
    color: "#2f6fed",
    kind: "missed",
    medium: "video",
    at: minutesAgo(40),
    conversationId: "aisha",
  },
  {
    id: "call-2",
    name: "Jordan Lee",
    color: "#1aa37a",
    kind: "outgoing",
    medium: "voice",
    at: minutesAgo(510),
    conversationId: "jordan",
  },
  {
    id: "call-3",
    name: "Maya Patel",
    color: "#e07a2f",
    kind: "incoming",
    medium: "voice",
    at: minutesAgo(1400),
    conversationId: "maya",
  },
  {
    id: "call-4",
    name: "Saturday studio",
    color: "#c45ad4",
    kind: "outgoing",
    medium: "video",
    at: minutesAgo(2600),
    conversationId: "studio",
  },
];

export const replies: Record<string, string[]> = {
  aisha: [
    "On it — I’ll send a tighter crop in a minute.",
    "Looks good. I’d lead with the neon sheet.",
    "Want me to rewrite the first line so it reads at pin size?",
  ],
  jordan: [
    "Sounds good. I’ll confirm with the printer.",
    "Proof is almost ready.",
    "Say the word and I’ll place the order.",
  ],
  maya: [
    "Okay. Text me when you’re close.",
    "All good on this end.",
    "I’ll keep the porch light on.",
  ],
  alex: [
    "Tracking should update tonight.",
    "I can reprint the label if you need it.",
    "Noted — I’ll hold the next batch.",
  ],
  sam: [
    "Appreciate it.",
    "I’ll try the A4 layout this afternoon.",
    "Thanks — that unblocked me.",
  ],
  studio: [
    "I’ll be there.",
    "Can someone grab extra clamps?",
    "Loft is open — see you soon.",
  ],
};

export const STORAGE_KEY = "relay-messenger.v1";
