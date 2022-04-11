let instance = null;

export class Router {
  constructor() {
    this._listeners = {};
  }
  static getInstance() {
    if (!instance) instance = new this();
    return instance;
  }
  use(eventId, listeners) {
    if (!this._listeners[eventId]) {
      this._listeners[eventId] = [listeners];
    } else {
      this._listeners[eventId].push(listeners);
    }
  }
  get(eventId) {
    return this._listeners[eventId];
  }
  run(eventId, options) {
    let listener = this.get(eventId);
    if (listener) {
      listener.forEach((handle) => {
        if (typeof handle === "function") handle(options);
      });
    }
  }
  dispatch(message) {
    let handlers = this.get(message.eventId);
    if (handlers) this.run(message.eventId, message);
  }
}
