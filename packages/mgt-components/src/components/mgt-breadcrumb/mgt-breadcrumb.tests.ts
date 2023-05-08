import '../../__mocks__/mock-media-match';
import { screen } from 'testing-library__dom';
import userEvent from '@testing-library/user-event';
import { fixture } from '@open-wc/testing-helpers';
import './mgt-breadcrumb';
import { BreadcrumbInfo, MgtBreadcrumb } from './mgt-breadcrumb';

describe('mgt-breadcrumb - tests', () => {
  it('should render', async () => {
    const component = await fixture(`
      <mgt-breadcrumb></mgt-breadcrumb>
    `);
    expect(component).toBeDefined();
  });
  it('should render with a single node', async () => {
    const component: MgtBreadcrumb = await fixture(`
      <mgt-breadcrumb></mgt-breadcrumb>
    `);
    component.breadcrumb = [
      {
        id: 'root-item',
        name: 'root-item'
      }
    ];
    await component.updateComplete;
    const listItems = await screen.findAllByRole('listitem');
    expect(listItems?.length).toBe(1);
    const root = await screen.findByText('root-item');
    expect(root).toBeDefined();
  });
  it('should render with three nodes', async () => {
    const component: MgtBreadcrumb = await fixture(`
      <mgt-breadcrumb></mgt-breadcrumb>
    `);
    component.breadcrumb = [
      {
        id: '0',
        name: 'root-item'
      },
      {
        id: '1',
        name: 'node-1'
      },
      {
        id: '2',
        name: 'node-2'
      }
    ];
    await component.updateComplete;
    const listItems = await screen.findAllByRole('listitem');
    expect(listItems?.length).toBe(3);
    const root = await screen.findByText('root-item');
    expect(root).toBeDefined();

    const buttons = await screen.findAllByRole('button');

    expect(buttons?.length).toBe(2);
  });

  it('should emit a clicked event when a button is clicked', async () => {
    const component: MgtBreadcrumb = await fixture(`
    <mgt-breadcrumb></mgt-breadcrumb>
  `);
    let eventEmitted;
    component.addEventListener('breadcrumbclick', (e: CustomEvent<BreadcrumbInfo>) => {
      eventEmitted = e.detail;
    });

    const rootNode = {
      id: '0',
      name: 'root-item'
    };
    component.breadcrumb = [
      rootNode,
      {
        id: '1',
        name: 'node-1'
      }
    ];
    await component.updateComplete;

    const buttons = await screen.findAllByRole('button');
    expect(buttons?.length).toBe(1);

    buttons[0].click();
    expect(eventEmitted).toBeDefined();
    expect(eventEmitted).toBe(rootNode);
  });

  it('should emit a clicked event when the enter button is pressed', async () => {
    const user = userEvent.setup();
    const component: MgtBreadcrumb = await fixture(`
    <mgt-breadcrumb></mgt-breadcrumb>
  `);
    let eventEmitted;
    component.addEventListener('breadcrumbclick', (e: CustomEvent<BreadcrumbInfo>) => {
      eventEmitted = e.detail;
    });

    const rootNode = {
      id: '0',
      name: 'root-item'
    };
    component.breadcrumb = [
      rootNode,
      {
        id: '1',
        name: 'node-1'
      }
    ];
    await component.updateComplete;

    const buttons = await screen.findAllByRole('button');
    expect(buttons?.length).toBe(1);
    buttons[0].focus();

    await user.keyboard('{Enter}');
    expect(eventEmitted).toBeDefined();
    expect(eventEmitted).toBe(rootNode);
  });

  it('should emit a clicked event when the space button is pressed', async () => {
    const user = userEvent.setup();
    const component: MgtBreadcrumb = await fixture(`
    <mgt-breadcrumb></mgt-breadcrumb>
  `);
    let eventEmitted;
    component.addEventListener('breadcrumbclick', (e: CustomEvent<BreadcrumbInfo>) => {
      eventEmitted = e.detail;
    });

    const rootNode = {
      id: '0',
      name: 'root-item'
    };
    component.breadcrumb = [
      rootNode,
      {
        id: '1',
        name: 'node-1'
      }
    ];
    await component.updateComplete;

    const buttons = await screen.findAllByRole('button');
    expect(buttons?.length).toBe(1);
    buttons[0].focus();

    await user.keyboard(' ');
    expect(eventEmitted).toBeDefined();
    expect(eventEmitted).toBe(rootNode);
  });

  it('should not emit a clicked event when a regular character button is pressed', async () => {
    const user = userEvent.setup();
    const component: MgtBreadcrumb = await fixture(`
    <mgt-breadcrumb></mgt-breadcrumb>
  `);
    let eventEmitted;
    component.addEventListener('breadcrumbclick', (e: CustomEvent<BreadcrumbInfo>) => {
      eventEmitted = e.detail;
    });

    const rootNode = {
      id: '0',
      name: 'root-item'
    };
    component.breadcrumb = [
      rootNode,
      {
        id: '1',
        name: 'node-1'
      }
    ];
    await component.updateComplete;

    const buttons = await screen.findAllByRole('button');
    expect(buttons?.length).toBe(1);
    buttons[0].focus();

    await user.keyboard('r');
    expect(eventEmitted).toBeUndefined();
  });

  it('should not emit a clicked event when the last node is clicked', async () => {
    const component: MgtBreadcrumb = await fixture(`
      <mgt-breadcrumb></mgt-breadcrumb>
    `);
    let eventEmitted;
    component.addEventListener('breadcrumbclick', (e: CustomEvent<BreadcrumbInfo>) => {
      eventEmitted = e.detail;
    });

    const rootNode = {
      id: '0',
      name: 'root-item'
    };
    component.breadcrumb = [
      rootNode,
      {
        id: '1',
        name: 'node-1'
      }
    ];
    await component.updateComplete;

    const notButton = await screen.findByText('node-1');
    notButton.click();

    expect(eventEmitted).toBeUndefined();
  });
});
