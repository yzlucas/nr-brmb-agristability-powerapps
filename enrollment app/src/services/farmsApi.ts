import type { IOperationResult } from '@microsoft/power-apps/data';
import { getClient } from '@microsoft/power-apps/data';
import { dataSourcesInfo } from '../../.power/schemas/appschemas/dataSourcesInfo';
import { FARMSAPIService } from '../generated/services/FARMSAPIService';

type FarmsApiResult<T = unknown> = IOperationResult<T>;

const FARMS_API_DATA_SOURCE_NAME = 'farms_20api_5fe39d1efd21a19d13_5f571039b465579741';
const GET_ENROLMENT_NOTICE_WORKFLOW_CALCULATION = 'GetEnrolmentNoticeWorkflowCalculation';

function ensureFarmsApiMetadata() {
  const sources = dataSourcesInfo as Record<string, { apis?: Record<string, unknown> } | undefined>;
  const farmsDataSource = sources[FARMS_API_DATA_SOURCE_NAME];
  if (!farmsDataSource) return;

  farmsDataSource.apis ??= {};
  farmsDataSource.apis[GET_ENROLMENT_NOTICE_WORKFLOW_CALCULATION] ??= {
    path: '/{connectionId}/calculations/enrolment-notice-workflow',
    method: 'GET',
    parameters: [
      { name: 'connectionId', in: 'path', required: true, type: 'string' },
      { name: 'participantPin', in: 'query', required: true, type: 'string' },
      { name: 'programYear', in: 'query', required: true, type: 'integer' },
    ],
    responseInfo: {
      default: {
        type: 'object',
      },
    },
  };
}

ensureFarmsApiMetadata();

const farmsApiClient = getClient(dataSourcesInfo);

const asFarmsResult = <T>(promise: Promise<IOperationResult<void>>): Promise<FarmsApiResult<T>> => {
  return promise as unknown as Promise<FarmsApiResult<T>>;
};

export const farmsApi = {
  getRoot: <T = unknown>() => asFarmsResult<T>(FARMSAPIService.GetRoot()),

  checkHealth: <T = unknown>(callstack = 'enrollment-app') => (
    asFarmsResult<T>(FARMSAPIService.GetCheckhealth(callstack))
  ),

  getAllCodeTables: <T = unknown>(effectiveAsOfDate?: string, codeTableName?: string) => (
    asFarmsResult<T>(FARMSAPIService.GetAllCodetables(effectiveAsOfDate, codeTableName))
  ),

  getOneCodeTable: <T = unknown>(codeTableName: string) => (
    asFarmsResult<T>(FARMSAPIService.GetOneCodetable(codeTableName))
  ),

  getOneCode: <T = unknown>(codeTableName: string, codeName: string) => (
    asFarmsResult<T>(FARMSAPIService.GetOneCode(codeTableName, codeName))
  ),

  getBenchmarkPerUnitsByProgramYear: <T = unknown>(programYear?: number) => (
    asFarmsResult<T>(FARMSAPIService.GetBenchmarkPerUnitsByProgramYear(programYear))
  ),

  getFairMarketValuesByProgramYear: <T = unknown>(programYear?: number) => (
    asFarmsResult<T>(FARMSAPIService.GetFairMarketValuesByProgramYear(programYear))
  ),

  getLineItemsByProgramYear: <T = unknown>(programYear?: number) => (
    asFarmsResult<T>(FARMSAPIService.GetLineItemsByProgramYear(programYear))
  ),

  getEnrolmentNoticeWorkflowCalculation: <T = unknown>(participantPin: string, programYear: number) => (
    farmsApiClient.executeAsync<{ participantPin: string; programYear: number }, T>({
      connectorOperation: {
        tableName: FARMS_API_DATA_SOURCE_NAME,
        operationName: GET_ENROLMENT_NOTICE_WORKFLOW_CALCULATION,
        parameters: {
          participantPin,
          programYear,
        },
      },
    })
  ),
};

export type { FarmsApiResult };
