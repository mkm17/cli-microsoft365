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
import command from './channel-add.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.CHANNEL_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
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
    loggerLogSpy = sinon.spy(logger, 'log');
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CHANNEL_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if both teamId and teamName options are passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        teamName: 'Team Name',
        name: 'Architecture Discussion',
        description: 'Architecture'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both channelId and channelName options are not passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        name: 'Architecture Discussion',
        description: 'Architecture'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the teamId is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        teamId: 'invalid GUID',
        name: 'Architecture Discussion',
        description: 'Architecture'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if unkown type is specified.', async () => {
    const actual = await command.validate({
      options: {
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        name: 'Architecture Discussion',
        type: 'invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if owner is not specified when creating private channel.', async () => {
    const actual = await command.validate({
      options: {
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        name: 'Architecture Discussion',
        type: 'private'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if owner is specified when not creating private channel.', async () => {
    const actual = await command.validate({
      options: {
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        name: 'Architecture Discussion',
        owner: 'John.Doe@contoso.com'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if owner is not specified when creating shared channel.', async () => {
    const actual = await command.validate({
      options: {
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        name: 'Architecture Discussion',
        type: 'shared'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if owner is specified when not creating a private or shared channel.', async () => {
    const actual = await command.validate({
      options: {
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        name: 'Architecture Discussion',
        owner: 'John.Doe@contoso.com'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('validates for a correct general channel input.', async () => {
    const actual = await command.validate({
      options: {
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        name: 'Architecture',
        description: 'Architecture meeting'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('validates for a correct private channel input.', async () => {
    const actual = await command.validate({
      options: {
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        name: 'Architecture',
        description: 'Architecture meeting',
        type: 'private',
        owner: 'john.doe@contoso.com'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('validates for a correct shared channel input.', async () => {
    const actual = await command.validate({
      options: {
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        name: 'Architecture',
        description: 'Architecture meeting',
        type: 'shared',
        owner: 'john.doe@contoso.com'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails to get team when team does not exists', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/me/joinedTeams') {
        return { value: [] };
      }

      throw 'The specified team does not exist in the Microsoft Teams';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        teamName: 'Team Name',
        name: 'Architecture Discussion',
        description: 'Architecture'
      }
    } as any), new CommandError('The specified team does not exist in the Microsoft Teams'));
  });

  it('fails when multiple teams with same name exists', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/me/joinedTeams') {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams",
          "@odata.count": 2,
          "value": [
            {
              "id": "00000000-0000-0000-0000-000000000000",
              "createdDateTime": null,
              "displayName": "Team Name",
              "description": "Team Description",
              "internalId": null,
              "classification": null,
              "specialization": null,
              "visibility": null,
              "webUrl": null,
              "isArchived": false,
              "isMembershipLimitedToOwners": null,
              "memberSettings": null,
              "guestSettings": null,
              "messagingSettings": null,
              "funSettings": null,
              "discoverySettings": null
            },
            {
              "id": "00000000-0000-0000-0000-000000000000",
              "createdDateTime": null,
              "displayName": "Team Name",
              "description": "Team Description",
              "internalId": null,
              "classification": null,
              "specialization": null,
              "visibility": null,
              "webUrl": null,
              "isArchived": false,
              "isMembershipLimitedToOwners": null,
              "memberSettings": null,
              "guestSettings": null,
              "messagingSettings": null,
              "funSettings": null,
              "discoverySettings": null
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        teamName: 'Team Name',
        name: 'Architecture Discussion',
        description: 'Architecture'
      }
    } as any), new CommandError('Multiple Microsoft Teams teams with name Team Name found. Found: 00000000-0000-0000-0000-000000000000.'));
  });

  it('handles selecting single result when multiple teams with the specified name found and cli is set to prompt', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/me/joinedTeams') {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams",
          "@odata.count": 2,
          "value": [
            {
              "id": "00000000-0000-0000-0000-000000000000",
              "createdDateTime": null,
              "displayName": "Team Name",
              "description": "Team Description",
              "internalId": null,
              "classification": null,
              "specialization": null,
              "visibility": null,
              "webUrl": null,
              "isArchived": false,
              "isMembershipLimitedToOwners": null,
              "memberSettings": null,
              "guestSettings": null,
              "messagingSettings": null,
              "funSettings": null,
              "discoverySettings": null
            },
            {
              "id": "00000000-0000-0000-0000-000000000000",
              "createdDateTime": null,
              "displayName": "Team Name",
              "description": "Team Description",
              "internalId": null,
              "classification": null,
              "specialization": null,
              "visibility": null,
              "webUrl": null,
              "isArchived": false,
              "isMembershipLimitedToOwners": null,
              "memberSettings": null,
              "guestSettings": null,
              "messagingSettings": null,
              "funSettings": null,
              "discoverySettings": null
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves({
      "id": "00000000-0000-0000-0000-000000000000",
      "createdDateTime": null,
      "displayName": "Team Name",
      "description": "Team Description",
      "internalId": null,
      "classification": null,
      "specialization": null,
      "visibility": null,
      "webUrl": null,
      "isArchived": false,
      "isMembershipLimitedToOwners": null,
      "memberSettings": null,
      "guestSettings": null,
      "messagingSettings": null,
      "funSettings": null,
      "discoverySettings": null
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/channels`) > -1) {
        return {
          "id": "19:d9c63a6d6a2644af960d74ea927bdfb0@thread.skype",
          "displayName": "Architecture Discussion",
          "description": null
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        teamName: 'Team Name',
        name: 'Architecture Discussion'
      }
    });

    assert(loggerLogSpy.calledWith({
      "id": "19:d9c63a6d6a2644af960d74ea927bdfb0@thread.skype",
      "displayName": "Architecture Discussion",
      "description": null
    }));
  });

  it('creates channel within the Microsoft Teams team in the tenant with description by team id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-28f0ac3a6402/channels`) {
        return {
          "id": "19:d9c63a6d6a2644af960d74ea927bdfb0@thread.skype",
          "displayName": "Architecture Discussion",
          "description": "Architecture"
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        name: 'Architecture Discussion',
        description: 'Architecture'
      }
    });

    assert(loggerLogSpy.calledWith({
      "id": "19:d9c63a6d6a2644af960d74ea927bdfb0@thread.skype",
      "displayName": "Architecture Discussion",
      "description": "Architecture"
    }));
  });

  it('creates channel within the Microsoft Teams team in the tenant without description by team id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-28f0ac3a6402/channels`) {
        return {
          "id": "19:d9c63a6d6a2644af960d74ea927bdfb0@thread.skype",
          "displayName": "Architecture Discussion",
          "description": null
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        name: 'Architecture Discussion'
      }
    });

    assert(loggerLogSpy.calledWith({
      "id": "19:d9c63a6d6a2644af960d74ea927bdfb0@thread.skype",
      "displayName": "Architecture Discussion",
      "description": null
    }));
  });

  it('creates private channel within the Microsoft Teams team by team id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-28f0ac3a6402/channels`) {
        return {
          "id": "19:d9c63a6d6a2644af960d74ea927bdfb0@thread.skype",
          "displayName": "Architecture Discussion",
          "membershipType": "private"
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        name: 'Architecture Discussion',
        type: 'private',
        owner: 'john.doe@contoso.com'
      }
    });

    assert(loggerLogSpy.calledWith({
      "id": "19:d9c63a6d6a2644af960d74ea927bdfb0@thread.skype",
      "displayName": "Architecture Discussion",
      "membershipType": "private"
    }));
  });

  it('creates shared channel within the Microsoft Teams team by team id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-28f0ac3a6402/channels`) {
        return {
          "id": "19:d9c63a6d6a2644af960d74ea927bdfb0@thread.skype",
          "displayName": "Architecture Discussion",
          "membershipType": "shared"
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        name: 'Architecture Discussion',
        type: 'shared',
        owner: 'john.doe@contoso.com'
      }
    });

    assert(loggerLogSpy.calledWith({
      "id": "19:d9c63a6d6a2644af960d74ea927bdfb0@thread.skype",
      "displayName": "Architecture Discussion",
      "membershipType": "shared"
    }));
  });

  it('creates channel within the Microsoft Teams team in the tenant by team name', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/me/joinedTeams') {
        return {
          "value": [
            {
              "id": "00000000-0000-0000-0000-000000000000",
              "createdDateTime": null,
              "displayName": "Team Name",
              "description": "Team Description",
              "internalId": null,
              "classification": null,
              "specialization": null,
              "visibility": null,
              "webUrl": null,
              "isArchived": false,
              "isMembershipLimitedToOwners": null,
              "memberSettings": null,
              "guestSettings": null,
              "messagingSettings": null,
              "funSettings": null,
              "discoverySettings": null
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/channels`) > -1) {
        return {
          "id": "19:d9c63a6d6a2644af960d74ea927bdfb0@thread.skype",
          "displayName": "Architecture Discussion",
          "description": null
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        teamName: 'Team Name',
        name: 'Architecture Discussion'
      }
    });

    assert(loggerLogSpy.calledWith({
      "id": "19:d9c63a6d6a2644af960d74ea927bdfb0@thread.skype",
      "displayName": "Architecture Discussion",
      "description": null
    }));
  });

  it('correctly handles error when adding a channel', async () => {
    const error = {
      "error": {
        "code": "UnknownError",
        "message": "An error has occurred",
        "innerError": {
          "date": "2022-02-14T13:27:37",
          "request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c",
          "client-request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c"
        }
      }
    };
    sinon.stub(request, 'post').rejects(error);

    await assert.rejects(command.action(logger, {
      options: {
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        name: 'Architecture Discussion'
      }
    } as any), new CommandError('An error has occurred'));
  });
});
