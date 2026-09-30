import { EventEmitter } from "events";

export interface MailUpdateEvent {
  userId: string;
  timestamp: string;
}

const emitter = new EventEmitter();

export function notifyMailUpdated(userId: string) {
  console.log(`📢 Emitting mailUpdated event for user ${userId}`);
  emitter.emit("mailUpdated", { userId, timestamp: new Date().toISOString() });
}

export function subscribeMailUpdates(listener: (event: MailUpdateEvent) => void) {
  emitter.on("mailUpdated", listener);
  return () => emitter.off("mailUpdated", listener);
}
