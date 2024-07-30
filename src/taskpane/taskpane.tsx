import * as React from "react";
import ReactDOM from "react-dom/client";

import App from "../components/App";


function render(conversationId: string) {
  const root = ReactDOM.createRoot(document.getElementById("root")!);
  root.render(<App key={conversationId} />);
}

Office.onReady(() => {
  render(Office.context.mailbox.item!.conversationId);
});
