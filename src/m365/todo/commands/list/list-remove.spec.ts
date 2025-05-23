import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './list-remove.js';
import { settingsNames } from '../../../../settingsNames.js';
import { accessToken } from '../../../../utils/accessToken.js';

describe(commands.LIST_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(accessToken, 'assertDelegatedAccessToken').returns();
    auth.connection.active = true;
    sinon.stub(cli, 'promptForConfirmation').resolves(true);
    commandInfo = cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.delete,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LIST_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('removes a To Do task list by name', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists?$filter=displayName eq 'FooList'`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#lists",
          "value": [
            {
              "@odata.etag": "W/\"m1fdwWoFiE2YS9yegTKoYwAA/ZGllw==\"",
              "displayName": "FooList",
              "isOwner": true,
              "isShared": false,
              "wellknownListName": "none",
              "id": "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIiAAA="
            }
          ]
        };
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists/AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIiAAA=`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        name: "FooList"
      }
    } as any);
    assert.strictEqual(log.length, 0);
  });

  it('removes a To Do task list by name when force option is passed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists?$filter=displayName eq 'FooList'`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#lists",
          "value": [
            {
              "@odata.etag": "W/\"m1fdwWoFiE2YS9yegTKoYwAA/ZGllw==\"",
              "displayName": "FooList",
              "isOwner": true,
              "isShared": false,
              "wellknownListName": "none",
              "id": "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIiAAA="
            }
          ]
        };
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists/AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIiAAA=`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        name: "FooList",
        force: true
      }
    } as any);
    assert.strictEqual(log.length, 0);
  });

  it('removes a To Do task list by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists?$filter=displayName eq 'FooList'`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#lists",
          "value": [
            {
              "@odata.etag": "W/\"m1fdwWoFiE2YS9yegTKoYwAA/ZGllw==\"",
              "displayName": "FooList",
              "isOwner": true,
              "isShared": false,
              "wellknownListName": "none",
              "id": "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIiAAA="
            }
          ]
        };
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists/AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIiAAA=`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        id: "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIiAAA="
      }
    } as any);
    assert.strictEqual(log.length, 0);
  });

  it('handles error correctly when a list is not found for a specific name', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists?$filter=displayName eq 'FooList'`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#lists",
          "value": []
        };
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'delete').callsFake(async () => {
      return;
    });
    await assert.rejects(command.action(logger, { options: { name: "FooList" } } as any), new CommandError('The list FooList cannot be found'));
  });

  it('handles error correctly', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists?$filter=displayName eq 'FooList'`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#lists",
          "value": [
            {
              "@odata.etag": "W/\"m1fdwWoFiE2YS9yegTKoYwAA/ZGllw==\"",
              "displayName": "FooList",
              "isOwner": true,
              "isShared": false,
              "wellknownListName": "none",
              "id": "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIiAAA="
            }
          ]
        };
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'delete').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: { name: "FooList" } } as any), new CommandError('An error has occurred'));
  });

  it('prompts before removing the list when force option not passed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists?$filter=displayName eq 'FooList'`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#lists",
          "value": [
            {
              "@odata.etag": "W/\"m1fdwWoFiE2YS9yegTKoYwAA/ZGllw==\"",
              "displayName": "FooList",
              "isOwner": true,
              "isShared": false,
              "wellknownListName": "none",
              "id": "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIiAAA="
            }
          ]
        };
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists/AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIiAAA=`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    const confirmationStub = sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, {
      options: {
        name: "FooList"
      }
    } as any);
    assert(confirmationStub.calledOnce);
  });

  it('fails validation if both name and id are not set', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        name: null,
        id: null
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when all parameters are valid', async () => {
    const actual = await command.validate({
      options: {
        name: 'Foo'
      }
    }, commandInfo);

    assert.strictEqual(actual, true);
  });

  it('fails validation if both name and id are set', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        name: 'foo',
        id: 'bar'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
});
