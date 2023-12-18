import React from 'react';
import { SampleChats, SampleMessagePreviewFromDexter } from './sampleData';
import { test, expect } from '@playwright/experimental-ct-react17';
import { ChatListItem } from './ChatListItem';
import { Chat } from '@microsoft/microsoft-graph-types';

test.describe('ChatListItem', () => {
  test.use({ viewport: { width: 500, height: 500 } });

  const chatSelected = (selected: Chat) => void 0;

  // group chat
  // - member added
  // - member removed
  // - "You: " prefix if you sent the latest message
  // - Member prefix if someone else sent the latest message
  // - title is the topic
  // - no topic, comma separated list of names
  test('groupchat: Other persons message', async ({ mount }) => {
    const component = await mount(
      <ChatListItem chat={SampleChats.group.SampleGroupChat} myId="id" onSelected={chatSelected} />
    );
    await expect(component).toContainText('Matthew: Hi everyone, call me Matt.');
  });

  test('groupchat: "You: " prefix on your message', async ({ mount }) => {
    let chat = SampleChats.group.SampleGroupChat;
    chat.lastMessagePreview = SampleMessagePreviewFromDexter;

    const component = await mount(
      <ChatListItem
        chat={SampleChats.group.SampleGroupChat}
        myId="a7530e0a-85f0-4f77-a617-2a5ab2a959ec"
        onSelected={chatSelected}
      />
    );
    component.textContent;
    await expect(component).toContainText('You: Note to self');
  });

  test('groupchat: implicit topic as a list of names', async ({ mount }) => {
    const component = await mount(
      <ChatListItem chat={SampleChats.group.SampleGroupChat} myId="id" onSelected={chatSelected} />
    );
    await expect(component).toContainText('Dexter, Andrew, Matthew, Mark, Luke, Johnathan');
  });

  test('groupchat: explicit topic', async ({ mount }) => {
    const component = await mount(
      <ChatListItem chat={SampleChats.group.SampleGroupChatMembershipChange} myId="id" onSelected={chatSelected} />
    );
    await expect(component).toContainText('Group Chat With New Members');
  });

  // one on one
  // - "You: " prefix if you sent the latest message
  // - No prefix if the other person sent the latest message
  // self chat
  test('oneonone: self chat', async ({ mount }) => {
    const component = await mount(
      <ChatListItem chat={SampleChats.oneOnOne.SampleSelfChat} myId="id" onSelected={chatSelected} />
    );
    await expect(component).toContainText('Note to self');
  });

  test('oneonone: other persons message', async ({ mount }) => {
    const component = await mount(
      <ChatListItem chat={SampleChats.oneOnOne.SampleChat} myId="id" onSelected={chatSelected} />
    );
    await expect(component).toContainText('Andrew: This is a sample message preview.');
  });

  test('oneonone: "You: " prefix if you sent the latest message', async ({ mount }) => {
    let chat = SampleChats.oneOnOne.SampleChat;
    chat.lastMessagePreview = SampleMessagePreviewFromDexter;

    const component = await mount(
      <ChatListItem chat={chat} myId="a7530e0a-85f0-4f77-a617-2a5ab2a959ec" onSelected={chatSelected} />
    );
    await expect(component).toContainText('You: Note to self');
  });

  // timestamp rendering tests
  test('timestamp: from today', async ({ mount }) => {
    const today = new Date();
    today.setHours(12, 0, 0, 0);

    const component = await mount(
      <ChatListItem chat={SampleChats.oneOnOne.SampleTodayChat} myId="id" onSelected={chatSelected} />
    );
    await expect(component).toContainText(
      today.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' }).toString()
    );
  });

  test('timestamp: from yesterday', async ({ mount }) => {
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    yesterday.setHours(12, 0, 0, 0);

    const component = await mount(
      <ChatListItem chat={SampleChats.oneOnOne.SampleYesterdayChat} myId="id" onSelected={chatSelected} />
    );
    await expect(component).toContainText(yesterday.toLocaleDateString([], { month: 'numeric', day: 'numeric' }));
  });
});
