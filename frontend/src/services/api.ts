import axios from 'axios';
import { ParseFileResponse, CreateDatabaseResponse, ConfigStatus } from '../types';

const API_BASE = '/api';

const apiClient = axios.create({
  baseURL: API_BASE,
});

export const setupAPI = {
  validateSQL: async (file: File): Promise<any> => {
    const formData = new FormData();
    formData.append('file', file);
    const response = await apiClient.post('/setup/validate-sql', formData);
    return response.data;
  },

  parseFile: async (file: File): Promise<ParseFileResponse> => {
    const formData = new FormData();
    formData.append('file', file);
    const response = await apiClient.post('/setup/parse-file', formData);
    return response.data;
  },

  createDatabase: async (file: File, mapping: any, recreate: boolean = true): Promise<CreateDatabaseResponse> => {
    const formData = new FormData();
    formData.append('file', file);
    formData.append('mapping_json', JSON.stringify(mapping));
    formData.append('recreate', recreate.toString());
    const response = await apiClient.post('/setup/create-database', formData);
    return response.data;
  },

  loadDatabase: async (): Promise<any> => {
    const response = await apiClient.get('/setup/load-database');
    return response.data;
  },
};

export const configAPI = {
  getStatus: async (): Promise<ConfigStatus> => {
    const response = await apiClient.get('/config/status');
    return response.data;
  },
};

export default apiClient;
