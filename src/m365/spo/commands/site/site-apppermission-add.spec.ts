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
import command from './site-apppermission-add.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.SITE_APPPERMISSION_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  //#region mocks
  const applicationMock = { "value": [{ "appId": "89ea5c94-7736-4e25-95ad-3fa95f62b66e", "displayName": "Foo App" }] };
  //#endregion

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
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
      request.patch,
      global.setTimeout,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SITE_APPPERMISSION_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation with an incorrect URL', async () => {
    const actual = await command.validate({
      options: {
        siteUrl: 'https;//contoso,sharepoint:com/sites/sitecollection-name',
        permission: "write",
        appId: "89ea5c94-7736-4e25-95ad-3fa95f62b66e"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the appId is not a valid GUID', async () => {
    const actual = await command.validate({
      options: {
        siteUrl: "https://contoso.sharepoint.com/sites/sitecollection-name",
        permission: "write",
        appId: "123"
      }
    }, commandInfo);

    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both appId and appDisplayName options are not specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        siteUrl: "https://contoso.sharepoint.com/sites/sitecollection-name",
        permission: "write"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation with a correct URL', async () => {
    const actual = await command.validate({
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        permission: "write",
        appId: "89ea5c94-7736-4e25-95ad-3fa95f62b66e"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if invalid value specified for permission', async () => {
    const actual = await command.validate({
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        permission: "Invalid",
        appId: "89ea5c94-7736-4e25-95ad-3fa95f62b66e"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails when passing a site that does not exist', async () => {
    const siteError = {
      "error": {
        "code": "itemNotFound",
        "message": "Requested site could not be found",
        "innerError": {
          "date": "2021-03-03T08:58:02",
          "request-id": "4e054f93-0eba-4743-be47-ce36b5f91120",
          "client-request-id": "dbd35b28-0ec3-6496-1279-0e1da3d028fe"
        }
      }
    };
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('non-existing') === -1) {
        return { value: [] };
      }
      throw siteError;
    });

    await assert.rejects(command.action(logger, {
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name-non-existing',
        permission: "write",
        appId: "89ea5c94-7736-4e25-95ad-3fa95f62b66e"
      }
    } as any), new CommandError('Requested site could not be found'));
  });

  it('fails to get Microsoft Entra app when Microsoft Entra app does not exists', async () => {
    const getRequestStub = sinon.stub(request, 'get');
    getRequestStub.onCall(0)
      .callsFake(async (opts) => {
        if ((opts.url as string).indexOf('/v1.0/sites/') > -1) {
          return {
            "id": "contoso.sharepoint.com,00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000",
            "displayName": "sitecollection-name",
            "name": "sitecollection-name",
            "createdDateTime": "2021-03-09T20:56:00Z",
            "lastModifiedDateTime": "2021-03-09T20:56:01Z",
            "webUrl": "https://contoso.sharepoint.com/sites/sitecollection-name"
          };
        }
        throw 'Invalid request';
      });

    getRequestStub.onCall(1)
      .callsFake(async (opts) => {
        if ((opts.url as string).indexOf('/v1.0/myorganization/servicePrincipals?$select=appId,displayName&$filter=') > -1) {
          return { value: [] };
        }
        throw 'The specified Microsoft Entra app does not exist';
      });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        permission: "write",
        appId: "89ea5c94-7736-4e25-95ad-3fa95f62b66e"
      }
    } as any), new CommandError('The specified Microsoft Entra app does not exist'));
  });

  it('fails when multiple Microsoft Entra apps with same name exists', async () => {
    const getRequestStub = sinon.stub(request, 'get');
    getRequestStub.onCall(0)
      .callsFake(async (opts) => {
        if ((opts.url as string).indexOf('/v1.0/sites/') > -1) {
          return {
            "id": "contoso.sharepoint.com,00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000",
            "displayName": "sitecollection-name",
            "name": "sitecollection-name",
            "createdDateTime": "2021-03-09T20:56:00Z",
            "lastModifiedDateTime": "2021-03-09T20:56:01Z",
            "webUrl": "https://contoso.sharepoint.com/sites/sitecollection-name"
          };
        }
        throw 'Invalid request';
      });

    getRequestStub.onCall(1)
      .callsFake(async (opts) => {
        if ((opts.url as string).indexOf('/v1.0/myorganization/servicePrincipals?$select=appId,displayName&') > -1) {
          return {
            "value": [
              {
                "appId": "3166f9d8-f4e9-4b56-b634-dafcc9ecba8e",
                "displayName": "Foo App"
              },
              {
                "appId": "9bd7b7c0-e4a7-4b85-b0c6-20aaca0e25b7",
                "displayName": "Foo App"
              }
            ]
          };
        }
        throw 'Multiple Microsoft Entra apps with displayName Foo App found: 3166f9d8-f4e9-4b56-b634-dafcc9ecba8e,9bd7b7c0-e4a7-4b85-b0c6-20aaca0e25b7';
      });

    await assert.rejects(command.action(logger, {
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        permission: "write",
        appDisplayName: "Foo App"
      }
    } as any), new CommandError('Multiple Microsoft Entra apps with displayName Foo App found: 3166f9d8-f4e9-4b56-b634-dafcc9ecba8e,9bd7b7c0-e4a7-4b85-b0c6-20aaca0e25b7'));
  });

  it('adds an application permission to the site by appId', async () => {
    const getRequestStub = sinon.stub(request, 'get');
    getRequestStub.onCall(0)
      .callsFake(async (opts) => {
        if ((opts.url as string).indexOf('/v1.0/sites/') > -1) {
          return {
            "id": "contoso.sharepoint.com,00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000",
            "displayName": "sitecollection-name",
            "name": "sitecollection-name",
            "createdDateTime": "2021-03-09T20:56:00Z",
            "lastModifiedDateTime": "2021-03-09T20:56:01Z",
            "webUrl": "https://contoso.sharepoint.com/sites/sitecollection-name"
          };
        }
        throw 'Invalid request';
      });

    getRequestStub.onCall(1)
      .callsFake(async (opts) => {
        if ((opts.url as string).indexOf('/v1.0/myorganization/servicePrincipals?$select=appId,displayName&') > -1) {
          return applicationMock;
        }

        throw 'Invalid request';
      });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/permissions') > -1) {
        return {
          "id": "aTowaS50fG1zLnNwLmV4dHxjY2EwMDE2OS1kMzhiLTQ2MmYtYTNiNC1mMzU2NmIxNjJmMmRAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0",
          "roles": [
            "write"
          ],
          "grantedToIdentities": [
            {
              "application": {
                "displayName": "Foo App",
                "id": "89ea5c94-7736-4e25-95ad-3fa95f62b66e"
              }
            }
          ]
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        siteUrl: "https://contoso.sharepoint.com/sites/sitecollection-name",
        permission: "write",
        appId: "89ea5c94-7736-4e25-95ad-3fa95f62b66e",
        output: "json"
      }
    });
    assert(loggerLogSpy.calledWith({
      "id": "aTowaS50fG1zLnNwLmV4dHxjY2EwMDE2OS1kMzhiLTQ2MmYtYTNiNC1mMzU2NmIxNjJmMmRAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0",
      "roles": [
        "write"
      ],
      "grantedToIdentities": [
        {
          "application": {
            "displayName": "Foo App",
            "id": "89ea5c94-7736-4e25-95ad-3fa95f62b66e"
          }
        }
      ]
    }));
  });

  it('adds an application permission to the site by appDisplayName', async () => {
    const getRequestStub = sinon.stub(request, 'get');
    getRequestStub.onCall(0)
      .callsFake(async (opts) => {
        if ((opts.url as string).indexOf('/v1.0/sites/') > -1) {
          return {
            "id": "contoso.sharepoint.com,00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000",
            "displayName": "sitecollection-name",
            "name": "sitecollection-name",
            "createdDateTime": "2021-03-09T20:56:00Z",
            "lastModifiedDateTime": "2021-03-09T20:56:01Z",
            "webUrl": "https://contoso.sharepoint.com/sites/sitecollection-name"
          };
        }
        throw 'Invalid request';
      });

    getRequestStub.onCall(1)
      .callsFake(async (opts) => {
        if ((opts.url as string).indexOf('/v1.0/myorganization/servicePrincipals?$select=appId,displayName&') > -1) {
          return applicationMock;
        }

        throw 'Invalid request';
      });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/permissions') > -1) {
        return {
          "id": "aTowaS50fG1zLnNwLmV4dHxjY2EwMDE2OS1kMzhiLTQ2MmYtYTNiNC1mMzU2NmIxNjJmMmRAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0",
          "roles": [
            "write"
          ],
          "grantedToIdentities": [
            {
              "application": {
                "displayName": "Foo App",
                "id": "89ea5c94-7736-4e25-95ad-3fa95f62b66e"
              }
            }
          ]
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        siteUrl: "https://contoso.sharepoint.com/sites/sitecollection-name",
        permission: "write",
        appDisplayName: "Foo App",
        output: "json"
      }
    });
    assert(loggerLogSpy.calledWith({
      "id": "aTowaS50fG1zLnNwLmV4dHxjY2EwMDE2OS1kMzhiLTQ2MmYtYTNiNC1mMzU2NmIxNjJmMmRAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0",
      "roles": [
        "write"
      ],
      "grantedToIdentities": [
        {
          "application": {
            "displayName": "Foo App",
            "id": "89ea5c94-7736-4e25-95ad-3fa95f62b66e"
          }
        }
      ]
    }));
  });

  it('adds an application permission to the site by appId and appDisplayName', async () => {
    const getRequestStub = sinon.stub(request, 'get');
    getRequestStub.onCall(0)
      .callsFake(async (opts) => {
        if ((opts.url as string).indexOf('/v1.0/sites/') > -1) {
          return {
            "id": "contoso.sharepoint.com,00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000",
            "displayName": "sitecollection-name",
            "name": "sitecollection-name",
            "createdDateTime": "2021-03-09T20:56:00Z",
            "lastModifiedDateTime": "2021-03-09T20:56:01Z",
            "webUrl": "https://contoso.sharepoint.com/sites/sitecollection-name"
          };
        }
        throw 'Invalid request';
      });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/permissions') > -1) {
        return {
          "id": "aTowaS50fG1zLnNwLmV4dHxjY2EwMDE2OS1kMzhiLTQ2MmYtYTNiNC1mMzU2NmIxNjJmMmRAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0",
          "roles": [
            "write"
          ],
          "grantedToIdentities": [
            {
              "application": {
                "displayName": "Foo App",
                "id": "89ea5c94-7736-4e25-95ad-3fa95f62b66e"
              }
            }
          ]
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        siteUrl: "https://contoso.sharepoint.com/sites/sitecollection-name",
        permission: "write",
        appId: "89ea5c94-7736-4e25-95ad-3fa95f62b66e",
        appDisplayName: "Foo App",
        output: "json"
      }
    });
    assert(loggerLogSpy.calledWith({
      "id": "aTowaS50fG1zLnNwLmV4dHxjY2EwMDE2OS1kMzhiLTQ2MmYtYTNiNC1mMzU2NmIxNjJmMmRAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0",
      "roles": [
        "write"
      ],
      "grantedToIdentities": [
        {
          "application": {
            "displayName": "Foo App",
            "id": "89ea5c94-7736-4e25-95ad-3fa95f62b66e"
          }
        }
      ]
    }));
  });

  it('adds an application permission to the site and elevates it to manage', async () => {
    const getRequestStub = sinon.stub(request, 'get');
    getRequestStub.onCall(0)
      .callsFake(async (opts) => {
        if ((opts.url as string).indexOf('/v1.0/sites/') > -1) {
          return {
            "id": "contoso.sharepoint.com,00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000",
            "displayName": "sitecollection-name",
            "name": "sitecollection-name",
            "createdDateTime": "2021-03-09T20:56:00Z",
            "lastModifiedDateTime": "2021-03-09T20:56:01Z",
            "webUrl": "https://contoso.sharepoint.com/sites/sitecollection-name"
          };
        }
        throw 'Invalid request';
      });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/permissions') > -1) {
        return {
          "id": "aTowaS50fG1zLnNwLmV4dHxjY2EwMDE2OS1kMzhiLTQ2MmYtYTNiNC1mMzU2NmIxNjJmMmRAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0",
          "roles": [
            "write"
          ],
          "grantedToIdentities": [
            {
              "application": {
                "displayName": "Foo App",
                "id": "89ea5c94-7736-4e25-95ad-3fa95f62b66e"
              }
            }
          ]
        };
      }
      throw 'Invalid request';
    });

    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/permissions/aTowaS5') > -1) {
        return {
          "id": "aTowaS50fG1zLnNwLmV4dHxjY2EwMDE2OS1kMzhiLTQ2MmYtYTNiNC1mMzU2NmIxNjJmMmRAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0",
          "roles": [
            "write"
          ],
          "grantedToIdentities": [
            {
              "application": {
                "id": "89ea5c94-7736-4e25-95ad-3fa95f62b66e"
              }
            }
          ]
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        siteUrl: "https://contoso.sharepoint.com/sites/sitecollection-name",
        permission: "manage",
        appId: "89ea5c94-7736-4e25-95ad-3fa95f62b66e",
        appDisplayName: "Foo App",
        output: "json"
      }
    });

    assert(loggerLogSpy.calledWith({
      "id": "aTowaS50fG1zLnNwLmV4dHxjY2EwMDE2OS1kMzhiLTQ2MmYtYTNiNC1mMzU2NmIxNjJmMmRAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0",
      "roles": [
        "write"
      ],
      "grantedToIdentities": [
        {
          "application": {
            "id": "89ea5c94-7736-4e25-95ad-3fa95f62b66e"
          }
        }
      ]
    }));
  });
});
