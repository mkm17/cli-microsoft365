import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './task-set.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.TASK_SET, () => {
  const taskResponse = {
    "planId": "8QZEH7b3wkS_bGQobscsM5gADCBb",
    "bucketId": "IK8tuFTwQEa5vTonM7ZMRZgAKdno",
    "title": "My Planner Task",
    "orderHint": "8585622710787367671",
    "assigneePriority": "",
    "percentComplete": 0,
    "startDateTime": null,
    "createdDateTime": "2021-12-12T19:03:26.7408136Z",
    "dueDateTime": null,
    "hasDescription": false,
    "previewType": "automatic",
    "completedDateTime": null,
    "completedBy": null,
    "referenceCount": 0,
    "checklistItemCount": 0,
    "activeChecklistItemCount": 0,
    "conversationThreadId": null,
    "id": "Z-RLQGfppU6H3663DBzfs5gAMD3o",
    "createdBy": {
      "user": {
        "displayName": null,
        "id": "dd8b99a7-77c6-4238-a609-396d27844921"
      }
    },
    "appliedCategories": {
      "category1": true,
      "category2": true
    },
    "assignments": {}
  };

  const taskResponseWithDetails = {
    "planId": "8QZEH7b3wkS_bGQobscsM5gADCBb",
    "bucketId": "IK8tuFTwQEa5vTonM7ZMRZgAKdno",
    "title": "My Planner Task",
    "orderHint": "8585622710787367671",
    "assigneePriority": "",
    "percentComplete": 0,
    "startDateTime": null,
    "createdDateTime": "2021-12-12T19:03:26.7408136Z",
    "dueDateTime": null,
    "hasDescription": false,
    "previewType": "automatic",
    "completedDateTime": null,
    "completedBy": null,
    "referenceCount": 0,
    "checklistItemCount": 0,
    "activeChecklistItemCount": 0,
    "conversationThreadId": null,
    "id": "Z-RLQGfppU6H3663DBzfs5gAMD3o",
    "createdBy": {
      "user": {
        "displayName": null,
        "id": "dd8b99a7-77c6-4238-a609-396d27844921"
      }
    },
    "appliedCategories": {
      "category1": true,
      "category2": true
    },
    "assignments": {},
    "description": "My Task Description",
    "references": {},
    "checklist": {}
  };

  const taskResponseWithAssignments = {
    "planId": "8QZEH7b3wkS_bGQobscsM5gADCBb",
    "bucketId": "IK8tuFTwQEa5vTonM7ZMRZgAKdno",
    "title": "My Planner Task",
    "orderHint": "8585622710787367671",
    "assigneePriority": "",
    "percentComplete": 0,
    "startDateTime": null,
    "createdDateTime": "2021-12-12T19:03:26.7408136Z",
    "dueDateTime": null,
    "hasDescription": false,
    "previewType": "automatic",
    "completedDateTime": null,
    "completedBy": null,
    "referenceCount": 0,
    "checklistItemCount": 0,
    "activeChecklistItemCount": 0,
    "conversationThreadId": null,
    "id": "Z-RLQGfppU6H3663DBzfs5gAMD3o",
    "createdBy": {
      "user": {
        "displayName": null,
        "id": "dd8b99a7-77c6-4238-a609-396d27844921"
      }
    },
    "appliedCategories": {
      "category1": true,
      "category2": true
    },
    "assignments": {
      "949b16c1-a032-453e-a8ae-89a52bfc1d8a": {
        "assignedDateTime": "2021-12-12T19:39:28.0946158Z",
        "orderHint": "8585622689774142174P}",
        "assignedBy": {
          "user": {
            "displayName": null,
            "id": "dd8b99a7-77c6-4238-a609-396d27844921"
          }
        }
      }
    }
  };

  const groupByDisplayNameResponse: any = {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#groups",
    "value": [
      {
        "id": "0d0402ee-970f-4951-90b5-2f24519d2e40"
      }
    ]
  };

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
    auth.connection.accessTokens[(command as any).resource] = {
      accessToken: 'abc',
      expiresOn: new Date()
    };
    commandInfo = cli.getCommandInfo(command);
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName: string, defaultValue: any) => {
      if (settingName === 'prompt') {
        return false;
      }

      return defaultValue;
    });
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
    (command as any).planId = undefined;
    (command as any).bucketId = undefined;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      request.patch,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TASK_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation when both bucketId and bucketName are specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
        bucketId: 'IK8tuFTwQEa5vTonM7ZMRZgAKdno',
        bucketName: 'My Bucket'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when bucketName is specified but not planId, planTitle, or rosterId', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
        bucketName: 'My Bucket'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when bucketName is specified but planId, planTitle, and rosterId are specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
        bucketName: 'My Bucket',
        planId: '8QZEH7b3wkS_bGQobscsM5gADCBb',
        planTitle: 'My Planner',
        rosterId: 'DjL5xiKO10qut8LQgztpKskABWna'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when planTitle is specified without ownerGroupId or ownerGroupName', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
        bucketName: 'My Bucket',
        planTitle: 'My Planner Plan'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when planTitle is specified with both ownerGroupId and ownerGroupName', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
        bucketName: 'My Bucket',
        planTitle: 'My Planner Plan',
        ownerGroupId: '0d0402ee-970f-4951-90b5-2f24519d2e40',
        ownerGroupName: 'My Planner Group'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the ownerGroupId is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
        bucketName: 'My Bucket',
        planTitle: 'My Planner Plan',
        ownerGroupId: 'not-c49b-4fd4-8223-28f0ac3a6402'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if startDateTime contains invalid format.', async () => {
    const actual = await command.validate({
      options: {
        id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
        startDateTime: '2021-99-99'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if dueDateTime contains invalid format.', async () => {
    const actual = await command.validate({
      options: {
        id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
        dueDateTime: '2021-99-99'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if percentComplete contains invalid format.', async () => {
    const actual = await command.validate({
      options: {
        id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
        percentComplete: 'Not A Number'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if percentComplete is not between 0 and 100.', async () => {
    const actual = await command.validate({
      options: {
        id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
        percentComplete: 599
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if assignedToUserIds contains invalid guid.', async () => {
    const actual = await command.validate({
      options: {
        id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
        assignedToUserIds: "2e42fe76-3f42-4884-b325-aefd7a905446,8d1ff29c-a6f4-4786-b316-test"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if assignedToUserNames contains invalid user principal name', async () => {
    const assignedToUserNames = ['john.doe@contoso.com', 'foo'];
    const actual = await command.validate({
      options: {
        id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
        assignedToUserNames: assignedToUserNames.join(',')
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when both assignedToUserIds and assignedToUserNames are specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
        assignedToUserIds: "2e42fe76-3f42-4884-b325-aefd7a905446,8d1ff29c-a6f4-4786-b316-eb6030e1a09e",
        assignedToUserNames: "Allan.Carroll@contoso.onmicrosoft.com,Ida.Stevens@contoso.onmicrosoft.com"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if incorrect appliedCategory is specified.', async () => {
    const actual = await command.validate({
      options: {
        id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
        appliedCategories: "category1,category9"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if priority lower than 0 is specified.', async () => {
    const actual = await command.validate({
      options: {
        id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
        priority: -1
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if priority higher than 10 is specified.', async () => {
    const actual = await command.validate({
      options: {
        id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
        priority: 11
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if priority is specified which is a number with decimals.', async () => {
    const actual = await command.validate({
      options: {
        id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
        priority: 5.6
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if unknown priority label is specified.', async () => {
    const actual = await command.validate({
      options: {
        id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
        priority: 'invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when valid options specified', async () => {
    const actual = await command.validate({
      options: {
        id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
        title: 'My Planner Task'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('correctly updates planner task with title', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter('Z-RLQGfppU6H3663DBzfs5gAMD3o')}`) {
        return taskResponse;
      }

      throw 'Invalid Request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter('Z-RLQGfppU6H3663DBzfs5gAMD3o')}` &&
        JSON.stringify(opts.headers) === JSON.stringify({
          'accept': 'application/json'
        })) {
        return { "@odata.etag": "TestEtag" };
      }

      throw 'Invalid Request';
    });

    const options: any = {
      id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
      title: 'My Planner Task'
    };

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(taskResponse));
  });

  it('uses correct value for urgent priority', async () => {
    const requestPatchStub = sinon.stub(request, 'patch').resolves(taskResponse);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter('Z-RLQGfppU6H3663DBzfs5gAMD3o')}`) {
        return { "@odata.etag": "TestEtag" };
      }

      throw 'Invalid Request';
    });

    const options: any = {
      id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
      priority: 'Urgent'
    };

    await command.action(logger, { options: options } as any);
    assert.strictEqual(requestPatchStub.lastCall.args[0].data.priority, 1);
  });

  it('uses correct value for important priority', async () => {
    const requestPatchStub = sinon.stub(request, 'patch').resolves(taskResponse);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter('Z-RLQGfppU6H3663DBzfs5gAMD3o')}`) {
        return { "@odata.etag": "TestEtag" };
      }

      throw 'Invalid Request';
    });

    const options: any = {
      id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
      priority: 'Important'
    };

    await command.action(logger, { options: options } as any);
    assert.strictEqual(requestPatchStub.lastCall.args[0].data.priority, 3);
  });

  it('uses correct value for medium priority', async () => {
    const requestPatchStub = sinon.stub(request, 'patch').resolves(taskResponse);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter('Z-RLQGfppU6H3663DBzfs5gAMD3o')}`) {
        return { "@odata.etag": "TestEtag" };
      }

      throw 'Invalid Request';
    });

    const options: any = {
      id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
      priority: 'Medium'
    };

    await command.action(logger, { options: options } as any);
    assert.strictEqual(requestPatchStub.lastCall.args[0].data.priority, 5);
  });

  it('uses correct value for low priority', async () => {
    const requestPatchStub = sinon.stub(request, 'patch').resolves(taskResponse);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter('Z-RLQGfppU6H3663DBzfs5gAMD3o')}`) {
        return { "@odata.etag": "TestEtag" };
      }

      throw 'Invalid Request';
    });

    const options: any = {
      id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
      priority: 'Low'
    };

    await command.action(logger, { options: options } as any);
    assert.strictEqual(requestPatchStub.lastCall.args[0].data.priority, 9);
  });

  it('correctly updates planner task to bucket with bucketName, planTitle, and ownerGroupName', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter('Z-RLQGfppU6H3663DBzfs5gAMD3o')}`) {
        return taskResponse;
      }

      throw 'Invalid Request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter('My Planner Group')}'`) {
        return groupByDisplayNameResponse;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/0d0402ee-970f-4951-90b5-2f24519d2e40/planner/plans?$select=id,title`) {
        return {
          value: [
            {
              "title": "My Planner Plan",
              "id": "8QZEH7b3wkS_bGQobscsM5gADCBb"
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${formatting.encodeQueryParameter('8QZEH7b3wkS_bGQobscsM5gADCBb')}/buckets?$select=id,name`) {
        return {
          value: [
            {
              "name": "My Planner Bucket",
              "id": "IK8tuFTwQEa5vTonM7ZMRZgAKdno"
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter('Z-RLQGfppU6H3663DBzfs5gAMD3o')}` &&
        JSON.stringify(opts.headers) === JSON.stringify({
          'accept': 'application/json'
        })) {
        return { "@odata.etag": "TestEtag" };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/0d0402ee-970f-4951-90b5-2f24519d2e40/planner/plans`) {
        return {
          value: [
            {
              "title": "My Planner Plan",
              "id": "8QZEH7b3wkS_bGQobscsM5gADCBb"
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter('My Planner Group')}'&$select=id`) {
        return groupByDisplayNameResponse;
      }

      throw 'Invalid Request';
    });

    const options: any = {
      id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
      bucketName: 'My Planner Bucket',
      planTitle: 'My Planner Plan',
      ownerGroupName: 'My Planner Group'
    };

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(taskResponse));
  });

  it('correctly updates planner task to bucket with bucketName, planTitle, and ownerGroupId', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter('Z-RLQGfppU6H3663DBzfs5gAMD3o')}`) {
        return taskResponse;
      }

      throw 'Invalid Request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${formatting.encodeQueryParameter('8QZEH7b3wkS_bGQobscsM5gADCBb')}/buckets?$select=id,name`) {
        return {
          value: [
            {
              "name": "My Planner Bucket",
              "id": "IK8tuFTwQEa5vTonM7ZMRZgAKdno"
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter('Z-RLQGfppU6H3663DBzfs5gAMD3o')}` &&
        JSON.stringify(opts.headers) === JSON.stringify({
          'accept': 'application/json'
        })) {
        return { "@odata.etag": "TestEtag" };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/0d0402ee-970f-4951-90b5-2f24519d2e40/planner/plans?$select=id,title`) {
        return {
          value: [
            {
              "title": "My Planner Plan",
              "id": "8QZEH7b3wkS_bGQobscsM5gADCBb"
            }
          ]
        };
      }

      throw 'Invalid Request';
    });

    const options: any = {
      id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
      bucketName: 'My Planner Bucket',
      planTitle: 'My Planner Plan',
      ownerGroupId: '0d0402ee-970f-4951-90b5-2f24519d2e40'
    };

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(taskResponse));
  });

  it('correctly updates planner task to bucket with bucketName and planId', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter('Z-RLQGfppU6H3663DBzfs5gAMD3o')}`) {
        return taskResponse;
      }

      throw 'Invalid Request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${formatting.encodeQueryParameter('8QZEH7b3wkS_bGQobscsM5gADCBb')}/buckets?$select=id,name`) {
        return {
          value: [
            {
              "name": "My Planner Bucket",
              "id": "IK8tuFTwQEa5vTonM7ZMRZgAKdno"
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter('Z-RLQGfppU6H3663DBzfs5gAMD3o')}` &&
        JSON.stringify(opts.headers) === JSON.stringify({
          'accept': 'application/json'
        })) {
        return { "@odata.etag": "TestEtag" };
      }

      throw 'Invalid Request';
    });

    const options: any = {
      id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
      bucketName: 'My Planner Bucket',
      planId: '8QZEH7b3wkS_bGQobscsM5gADCBb'
    };

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(taskResponse));
  });

  it('correctly updates planner task with percentComplete by rosterId', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter('Z-RLQGfppU6H3663DBzfs5gAMD3o')}`) {
        return taskResponse;
      }

      throw 'Invalid Request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/planner/rosters/DjL5xiKO10qut8LQgztpKskABWna/plans?$select=id`) {
        return {
          "value": [{
            "id": '8QZEH7b3wkS_bGQobscsM5gADCBb'
          }]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/8QZEH7b3wkS_bGQobscsM5gADCBb/buckets?$select=id,name`) {
        return {
          "value": [
            {
              "name": "My Planner Bucket",
              "id": "IK8tuFTwQEa5vTonM7ZMRZgAKdno"
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter('Z-RLQGfppU6H3663DBzfs5gAMD3o')}` &&
        JSON.stringify(opts.headers) === JSON.stringify({
          'accept': 'application/json'
        })) {
        return {
          "@odata.etag": "TestEtag"
        };
      }

      throw 'Invalid request';
    });

    const options: any = {
      id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
      percentComplete: '50',
      rosterId: 'DjL5xiKO10qut8LQgztpKskABWna',
      bucketName: 'My Planner Bucket'
    };

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(taskResponse));
  });


  it('correctly updates planner task with assignedToUserIds', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter('Z-RLQGfppU6H3663DBzfs5gAMD3o')}`) {
        return taskResponseWithAssignments;
      }

      throw 'Invalid Request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter('Z-RLQGfppU6H3663DBzfs5gAMD3o')}` &&
        JSON.stringify(opts.headers) === JSON.stringify({
          'accept': 'application/json'
        })) {
        return { "@odata.etag": "TestEtag" };
      }

      throw 'Invalid Request';
    });

    const options: any = {
      id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
      assignedToUserIds: '949b16c1-a032-453e-a8ae-89a52bfc1d8a'
    };

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(taskResponseWithAssignments));
  });

  it('correctly updates planner task with assignedToUserNames', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter('Z-RLQGfppU6H3663DBzfs5gAMD3o')}`) {
        return taskResponseWithAssignments;
      }

      throw 'Invalid Request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '${formatting.encodeQueryParameter('user@contoso.onmicrosoft.com')}'&$select=id,userPrincipalName`) {
        return {
          value: [
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8a',
              userPrincipalName: 'user@contoso.onmicrosoft.com'
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter('Z-RLQGfppU6H3663DBzfs5gAMD3o')}` &&
        JSON.stringify(opts.headers) === JSON.stringify({
          'accept': 'application/json'
        })) {
        return { "@odata.etag": "TestEtag" };
      }

      throw 'Invalid Request';
    });

    const options: any = {
      id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
      assignedToUserNames: 'user@contoso.onmicrosoft.com'
    };

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(taskResponseWithAssignments));
  });

  it('correctly updates planner task with description', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter('Z-RLQGfppU6H3663DBzfs5gAMD3o')}/details` &&
        JSON.stringify(opts.headers) === JSON.stringify({
          'accept': 'application/json'
        })) {
        return { "@odata.etag": "TestEtag" };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter('Z-RLQGfppU6H3663DBzfs5gAMD3o')}` &&
        JSON.stringify(opts.headers) === JSON.stringify({
          'accept': 'application/json'
        })) {
        return { "@odata.etag": "TestEtag" };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks`) {
        return taskResponseWithDetails;
      }

      throw 'Invalid Request';
    });

    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter('Z-RLQGfppU6H3663DBzfs5gAMD3o')}/details`) {
        return {
          "description": "My Task Description",
          "references": {},
          "checklist": {}
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter('Z-RLQGfppU6H3663DBzfs5gAMD3o')}`) {
        return taskResponseWithDetails;
      }

      throw 'Invalid Request';
    });

    const options: any = {
      id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
      description: 'My Task Description'
    };

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(taskResponseWithDetails));
  });

  it('correctly updates planner task with appliedCategories, bucketId, startDateTime, dueDateTime, percentComplete, assigneePriority, orderHint, and priority', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter('Z-RLQGfppU6H3663DBzfs5gAMD3o')}`) {
        return taskResponse;
      }

      throw 'Invalid Request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter('Z-RLQGfppU6H3663DBzfs5gAMD3o')}` &&
        JSON.stringify(opts.headers) === JSON.stringify({
          'accept': 'application/json'
        })) {
        return { "@odata.etag": "TestEtag" };
      }

      throw 'Invalid Request';
    });

    const options: any = {
      id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
      appliedCategories: 'category1,category2',
      bucketId: 'IK8tuFTwQEa5vTonM7ZMRZgAKdno',
      startDateTime: '2014-01-01T00:00:00Z',
      dueDateTime: '2023-01-01T00:00:00Z',
      percentComplete: '50',
      assigneePriority: ' !',
      orderHint: ' !',
      priority: 3
    };

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(taskResponse));
  });

  it('fails when no bucket is found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${formatting.encodeQueryParameter('8QZEH7b3wkS_bGQobscsM5gADCBb')}/buckets?$select=id,name`) {
        return { value: [] };
      }

      throw 'Invalid Request';
    });

    const options: any = {
      id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
      bucketName: 'My Planner Bucket',
      planId: '8QZEH7b3wkS_bGQobscsM5gADCBb'
    };

    await assert.rejects(command.action(logger, { options: options } as any), new CommandError(`The specified bucket 'My Planner Bucket' does not exist.`));
  });

  it('fails when an invalid user is specified', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user%40contoso.onmicrosoft.com'&$select=id,userPrincipalName`) {
        return {
          value: [
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8a',
              userPrincipalName: 'user@contoso.onmicrosoft.com'
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user2%40contoso.onmicrosoft.com'&$select=id,userPrincipalName`) {
        return { value: [] };
      }

      throw 'Invalid Request';
    });

    const options: any = {
      id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
      assignedToUserNames: 'user@contoso.onmicrosoft.com,user2@contoso.onmicrosoft.com'
    };

    await assert.rejects(command.action(logger, { options: options } as any)
      , new CommandError('Cannot proceed with planner task update. The following users provided are invalid : user2@contoso.onmicrosoft.com'));
  });

  it('fails validation when ownerGroupName not found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/groups?$filter=displayName') > -1) {
        return { value: [] };
      }
      throw 'Invalid Request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
        bucketName: 'My Planner Bucket',
        planTitle: 'My Planner Plan',
        ownerGroupName: 'foo'
      }
    }), new CommandError(`The specified group 'foo' does not exist.`));
  });

  it('fails validation when task endpoint fails', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter('Z-RLQGfppU6H3663DBzfs5gAMD3o')}`) {
        return taskResponse;
      }

      throw 'Invalid Request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter('Z-RLQGfppU6H3663DBzfs5gAMD3o')}/details`) {
        throw 'Error fetching task';
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter('Z-RLQGfppU6H3663DBzfs5gAMD3o')}` &&
        JSON.stringify(opts.headers) === JSON.stringify({
          'accept': 'application/json'
        })) {
        throw 'Error fetching task';
      }

      throw 'Invalid Request';
    });


    await assert.rejects(command.action(logger, {
      options: {
        id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
        title: 'My Planner Task'
      }
    }), new CommandError(`Error fetching task`));
  });

  it('fails validation when task details endpoint fails', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter('Z-RLQGfppU6H3663DBzfs5gAMD3o')}`) {
        return taskResponse;
      }

      throw 'Invalid Request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter('Z-RLQGfppU6H3663DBzfs5gAMD3o')}/details`) {
        throw 'Error fetching task details';
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter('Z-RLQGfppU6H3663DBzfs5gAMD3o')}` &&
        JSON.stringify(opts.headers) === JSON.stringify({
          'accept': 'application/json'
        })) {
        return { "@odata.etag": "TestEtag" };
      }

      throw 'Invalid Request';
    });


    await assert.rejects(command.action(logger, {
      options: {
        id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
        description: 'My Task Description'
      }
    }), new CommandError(`Error fetching task details`));
  });

  it('correctly handles random API error', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });
});