import * as React from "react";

interface MessageListProps {
  messages: any[]; // Replace 'any' with the actual type of your messages
}

const MessageList: React.FC<MessageListProps> = ({ messages }) => {
  return (
    <div>
      {messages.map((item) => (
        <ul key={item.id}>
          <li>
            <span className="ms-font-l">{item.subject}</span>
          </li>
        </ul>
      ))}
    </div>
  );
};

export default MessageList;
