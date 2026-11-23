export type Tab = "chats" | "calls" | "people" | "settings";

export type Message = {
  id: string;
  senderId: string;
  text: string;
  sentAt: number;
};

export type Conversation = {
  id: string;
  title: string;
  subtitle: string;
  participants: string[];
  color: string;
  messages: Message[];
  unread: number;
  muted?: boolean;
  isGroup?: boolean;
};

export type CallRecord = {
  id: string;
  name: string;
  color: string;
  kind: "incoming" | "outgoing" | "missed";
  medium: "voice" | "video";
  at: number;
  conversationId: string;
};

export type Contact = {
  id: string;
  name: string;
  color: string;
  handle: string;
  conversationId?: string;
};
