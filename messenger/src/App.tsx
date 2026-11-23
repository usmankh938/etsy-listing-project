import { useEffect, useMemo, useRef, useState, type RefObject } from "react";
import {
  ME,
  STORAGE_KEY,
  replies,
  seedCalls,
  seedContacts,
  seedConversations,
} from "./data";
import type { Conversation, Message, Tab } from "./types";

function loadConversations(): Conversation[] {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return seedConversations;
    const parsed = JSON.parse(raw) as Conversation[];
    if (!Array.isArray(parsed) || parsed.length === 0) return seedConversations;
    return parsed;
  } catch {
    return seedConversations;
  }
}

function initials(name: string) {
  return name
    .split(" ")
    .filter(Boolean)
    .slice(0, 2)
    .map((part) => part[0]?.toUpperCase() ?? "")
    .join("");
}

function formatTime(ts: number) {
  const diff = Date.now() - ts;
  const minute = 60_000;
  const hour = 60 * minute;
  const day = 24 * hour;
  if (diff < minute) return "now";
  if (diff < hour) return `${Math.floor(diff / minute)}m`;
  if (diff < day) return `${Math.floor(diff / hour)}h`;
  return new Date(ts).toLocaleDateString(undefined, { month: "short", day: "numeric" });
}

function formatClock(ts: number) {
  return new Date(ts).toLocaleTimeString(undefined, {
    hour: "numeric",
    minute: "2-digit",
  });
}

function lastMessage(conversation: Conversation) {
  return conversation.messages[conversation.messages.length - 1];
}

function senderName(conversation: Conversation, senderId: string) {
  if (senderId === ME.id) return "You";
  const contact = seedContacts.find((item) => item.id === senderId);
  if (contact) return contact.name.split(" ")[0] ?? contact.name;
  return conversation.title;
}

function uid() {
  return `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

export default function App() {
  const [tab, setTab] = useState<Tab>("chats");
  const [query, setQuery] = useState("");
  const [conversations, setConversations] = useState<Conversation[]>(loadConversations);
  const [activeId, setActiveId] = useState<string | null>(null);
  const [draft, setDraft] = useState("");
  const [typing, setTyping] = useState(false);
  const [now, setNow] = useState(() => new Date());
  const threadRef = useRef<HTMLDivElement>(null);
  const replyTimer = useRef<number | null>(null);

  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(conversations));
  }, [conversations]);

  useEffect(() => {
    const tick = window.setInterval(() => setNow(new Date()), 30_000);
    return () => window.clearInterval(tick);
  }, []);

  useEffect(() => {
    return () => {
      if (replyTimer.current) window.clearTimeout(replyTimer.current);
    };
  }, []);

  useEffect(() => {
    if (threadRef.current) {
      threadRef.current.scrollTop = threadRef.current.scrollHeight;
    }
  }, [activeId, conversations, typing]);

  const active = conversations.find((item) => item.id === activeId) ?? null;
  const unreadTotal = conversations.reduce((sum, item) => sum + item.unread, 0);

  const filtered = useMemo(() => {
    const needle = query.trim().toLowerCase();
    const list = [...conversations].sort((a, b) => {
      const aTime = lastMessage(a)?.sentAt ?? 0;
      const bTime = lastMessage(b)?.sentAt ?? 0;
      return bTime - aTime;
    });
    if (!needle) return list;
    return list.filter((item) => {
      const haystack = [
        item.title,
        item.subtitle,
        ...item.messages.map((message) => message.text),
      ]
        .join(" ")
        .toLowerCase();
      return haystack.includes(needle);
    });
  }, [conversations, query]);

  function openChat(id: string) {
    setActiveId(id);
    setTab("chats");
    setConversations((current) =>
      current.map((item) => (item.id === id ? { ...item, unread: 0 } : item)),
    );
  }

  function sendMessage() {
    const text = draft.trim();
    if (!text || !active) return;
    const outgoing: Message = {
      id: uid(),
      senderId: ME.id,
      text,
      sentAt: Date.now(),
    };
    setDraft("");
    setConversations((current) =>
      current.map((item) =>
        item.id === active.id ? { ...item, messages: [...item.messages, outgoing] } : item,
      ),
    );

    const replyPool = replies[active.id] ?? replies[active.participants[0] ?? ""] ?? [];
    if (replyPool.length === 0) return;
    setTyping(true);
    if (replyTimer.current) window.clearTimeout(replyTimer.current);
    replyTimer.current = window.setTimeout(() => {
      const reply: Message = {
        id: uid(),
        senderId: active.participants[0] ?? active.id,
        text: replyPool[Math.floor(Math.random() * replyPool.length)] ?? "Got it.",
        sentAt: Date.now(),
      };
      setConversations((current) =>
        current.map((item) =>
          item.id === active.id ? { ...item, messages: [...item.messages, reply] } : item,
        ),
      );
      setTyping(false);
    }, 900);
  }

  function resetSampleData() {
    localStorage.removeItem(STORAGE_KEY);
    setConversations(seedConversations);
    setActiveId(null);
    setDraft("");
    setTyping(false);
    setQuery("");
    setTab("chats");
  }

  const clock = now.toLocaleTimeString(undefined, {
    hour: "numeric",
    minute: "2-digit",
  });

  return (
    <div className="stage">
      <div className="phone" data-testid="phone">
        <div className="status-bar">
          <span>{clock}</span>
          <span className="notch" aria-hidden="true" />
          <span className="status-icons">LTE · 84%</span>
        </div>

        {active ? (
          <ChatScreen
            conversation={active}
            draft={draft}
            typing={typing}
            threadRef={threadRef}
            onBack={() => {
              setActiveId(null);
              setTyping(false);
            }}
            onDraft={setDraft}
            onSend={sendMessage}
          />
        ) : (
          <>
            {tab === "chats" && (
              <ChatsScreen
                conversations={filtered}
                query={query}
                unreadTotal={unreadTotal}
                onQuery={setQuery}
                onOpen={openChat}
              />
            )}
            {tab === "calls" && <CallsScreen onOpen={openChat} />}
            {tab === "people" && <PeopleScreen onOpen={openChat} />}
            {tab === "settings" && <SettingsScreen onReset={resetSampleData} />}
            <nav className="tab-bar" aria-label="App tabs">
              <TabButton
                id="chats"
                label="Chats"
                active={tab === "chats"}
                badge={unreadTotal}
                onClick={() => setTab("chats")}
              />
              <TabButton id="calls" label="Calls" active={tab === "calls"} onClick={() => setTab("calls")} />
              <TabButton id="people" label="People" active={tab === "people"} onClick={() => setTab("people")} />
              <TabButton
                id="settings"
                label="Settings"
                active={tab === "settings"}
                onClick={() => setTab("settings")}
              />
            </nav>
          </>
        )}
      </div>
    </div>
  );
}

function TabButton({
  id,
  label,
  active,
  badge = 0,
  onClick,
}: {
  id: Tab;
  label: string;
  active: boolean;
  badge?: number;
  onClick: () => void;
}) {
  return (
    <button type="button" className={active ? "tab active" : "tab"} onClick={onClick} aria-current={active}>
      <span className={`glyph glyph-${id}`} aria-hidden="true" />
      <span>{label}</span>
      {badge > 0 ? <span className="tab-badge">{badge}</span> : null}
    </button>
  );
}

function ChatsScreen({
  conversations,
  query,
  unreadTotal,
  onQuery,
  onOpen,
}: {
  conversations: Conversation[];
  query: string;
  unreadTotal: number;
  onQuery: (value: string) => void;
  onOpen: (id: string) => void;
}) {
  return (
    <section className="screen">
      <header className="top">
        <p className="kicker">Relay</p>
        <h1>Chats</h1>
        {unreadTotal > 0 ? <p className="hint">{unreadTotal} unread</p> : <p className="hint">Sample inbox</p>}
      </header>
      <label className="search">
        <span className="visually-hidden">Search chats</span>
        <input
          value={query}
          onChange={(event) => onQuery(event.target.value)}
          placeholder="Search"
          autoComplete="off"
        />
      </label>
      <ul className="list">
        {conversations.length === 0 ? (
          <li className="empty">No chats match that search.</li>
        ) : null}
        {conversations.map((item) => {
          const last = lastMessage(item);
          const preview =
            last?.senderId === ME.id ? `You: ${last.text}` : (last?.text ?? "No messages yet");
          return (
            <li key={item.id}>
              <button type="button" className="row" onClick={() => onOpen(item.id)}>
                <span className="avatar" style={{ background: item.color }}>
                  {initials(item.title)}
                </span>
                <span className="row-body">
                  <span className="row-top">
                    <strong>{item.title}</strong>
                    <time>{last ? formatTime(last.sentAt) : ""}</time>
                  </span>
                  <span className={item.unread ? "preview unread" : "preview"}>{preview}</span>
                </span>
                {item.unread > 0 ? <span className="badge">{item.unread}</span> : null}
              </button>
            </li>
          );
        })}
      </ul>
    </section>
  );
}

function ChatScreen({
  conversation,
  draft,
  typing,
  threadRef,
  onBack,
  onDraft,
  onSend,
}: {
  conversation: Conversation;
  draft: string;
  typing: boolean;
  threadRef: RefObject<HTMLDivElement | null>;
  onBack: () => void;
  onDraft: (value: string) => void;
  onSend: () => void;
}) {
  return (
    <section className="screen thread-screen">
      <header className="chat-header">
        <button type="button" className="back" onClick={onBack} aria-label="Back to chats">
          ‹
        </button>
        <span className="avatar small" style={{ background: conversation.color }}>
          {initials(conversation.title)}
        </span>
        <div>
          <h1>{conversation.title}</h1>
          <p>{typing ? "Typing…" : conversation.subtitle}</p>
        </div>
      </header>
      <div className="thread" ref={threadRef}>
        {conversation.messages.map((message, index) => {
          const mine = message.senderId === ME.id;
          const previous = conversation.messages[index - 1];
          const showName =
            Boolean(conversation.isGroup) &&
            !mine &&
            previous?.senderId !== message.senderId;
          return (
            <div key={message.id} className={mine ? "bubble-wrap mine" : "bubble-wrap"}>
              {showName ? <span className="sender">{senderName(conversation, message.senderId)}</span> : null}
              <div className={mine ? "bubble mine" : "bubble"}>{message.text}</div>
              <time>{formatClock(message.sentAt)}</time>
            </div>
          );
        })}
        {typing ? (
          <div className="bubble-wrap">
            <div className="bubble typing" aria-label="Typing">
              <span />
              <span />
              <span />
            </div>
          </div>
        ) : null}
      </div>
      <form
        className="composer"
        onSubmit={(event) => {
          event.preventDefault();
          onSend();
        }}
      >
        <label className="visually-hidden" htmlFor="composer">
          Message
        </label>
        <input
          id="composer"
          value={draft}
          onChange={(event) => onDraft(event.target.value)}
          placeholder="Message"
          autoComplete="off"
        />
        <button type="submit" disabled={!draft.trim()}>
          Send
        </button>
      </form>
    </section>
  );
}

function CallsScreen({ onOpen }: { onOpen: (id: string) => void }) {
  return (
    <section className="screen">
      <header className="top">
        <p className="kicker">Relay</p>
        <h1>Calls</h1>
        <p className="hint">Recent · sample</p>
      </header>
      <ul className="list">
        {seedCalls.map((item) => (
          <li key={item.id}>
            <button type="button" className="row" onClick={() => onOpen(item.conversationId)}>
              <span className="avatar" style={{ background: item.color }}>
                {initials(item.name)}
              </span>
              <span className="row-body">
                <span className="row-top">
                  <strong className={item.kind === "missed" ? "missed" : undefined}>{item.name}</strong>
                  <time>{formatTime(item.at)}</time>
                </span>
                <span className="preview">
                  {item.kind} {item.medium}
                </span>
              </span>
            </button>
          </li>
        ))}
      </ul>
    </section>
  );
}

function PeopleScreen({ onOpen }: { onOpen: (id: string) => void }) {
  return (
    <section className="screen">
      <header className="top">
        <p className="kicker">Relay</p>
        <h1>People</h1>
        <p className="hint">{seedContacts.length} contacts</p>
      </header>
      <ul className="list">
        {seedContacts.map((item) => (
          <li key={item.id}>
            <button type="button" className="row" onClick={() => item.conversationId && onOpen(item.conversationId)}>
              <span className="avatar" style={{ background: item.color }}>
                {initials(item.name)}
              </span>
              <span className="row-body">
                <span className="row-top">
                  <strong>{item.name}</strong>
                </span>
                <span className="preview">{item.handle}</span>
              </span>
            </button>
          </li>
        ))}
      </ul>
    </section>
  );
}

function SettingsScreen({ onReset }: { onReset: () => void }) {
  return (
    <section className="screen settings">
      <header className="top">
        <p className="kicker">Relay</p>
        <h1>Settings</h1>
      </header>
      <div className="profile-card">
        <span className="avatar large" style={{ background: "#1aa37a" }}>
          {initials(ME.name)}
        </span>
        <div>
          <strong>{ME.name}</strong>
          <p>{ME.handle}</p>
        </div>
      </div>
      <p className="sample-note">
        This inbox is sample data so you can tap around immediately. Messages you send stay in this
        browser.
      </p>
      <button type="button" className="reset" onClick={onReset}>
        Reset sample chats
      </button>
    </section>
  );
}
