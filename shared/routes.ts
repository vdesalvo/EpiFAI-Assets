
import { z } from 'zod';

export const api = {
  manifest: {
    get: {
      method: 'GET' as const,
      path: '/manifest.xml' as const,
      responses: {
        200: z.string(),
      },
    },
  },
};
