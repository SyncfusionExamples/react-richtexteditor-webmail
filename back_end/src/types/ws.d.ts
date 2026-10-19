declare module "ws" {
  import { EventEmitter } from "events";

  class WebSocket extends EventEmitter {
    static readonly OPEN: number;
    readyState: number;
    send(data: string | ArrayBufferLike | Blob | ArrayBufferView): void;
    close(code?: number, reason?: string): void;
    terminate(): void;
  }

  export default WebSocket;
}
