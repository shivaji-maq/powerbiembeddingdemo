// hooks/usePersonalization/index.ts
import { useState, useCallback } from "react";
import axios from "axios";

const viteEnv = ((import.meta as any)?.env || {}) as Record<string, any>;

const personalizationApiBase =
  (viteEnv.VITE_PERSONALIZATION_API_BASE_URL as string | undefined)?.trim() ||
  "";
const hasPersonalizationApiBase = personalizationApiBase.length > 0;
const usePersonalizationApiInDev =
  String(viteEnv.VITE_USE_PERSONALIZATION_API || "").toLowerCase() ===
  "true";
const shouldUsePersonalizationApiInDev =
  usePersonalizationApiInDev && hasPersonalizationApiBase;
const localOnlyPersonalization =
  Boolean(viteEnv.DEV) && !shouldUsePersonalizationApiInDev;
let personalizationApiUnavailable = localOnlyPersonalization || !hasPersonalizationApiBase;

const withApiBase = (path: string) =>
  personalizationApiBase ? `${personalizationApiBase}${path}` : path;

export interface PersonalizationData {
  id?: string;
  userId: string;
  reportId: string;
  workspaceId: string;
  filtersJson: string;
  bookmarksJson: string;
  activePage: string;
  settingsJson: string;
  updatedAt?: string;
}

const getLocalPersonalizationKey = (userId: string, reportId: string) =>
  `pbi_personalization_${userId}_${reportId}`;

const readLocalPersonalization = (userId: string, reportId: string) => {
  try {
    const stored = window.localStorage.getItem(
      getLocalPersonalizationKey(userId, reportId)
    );
    return stored ? JSON.parse(stored) : null;
  } catch (error) {
    console.warn("Failed to parse local personalization", error);
    return null;
  }
};

const writeLocalPersonalization = (data: Partial<PersonalizationData>) => {
  if (!data.userId || !data.reportId) {
    return null;
  }

  const payload = {
    ...data,
    updatedAt: new Date().toISOString(),
  };

  window.localStorage.setItem(
    getLocalPersonalizationKey(data.userId, data.reportId),
    JSON.stringify(payload)
  );

  return payload;
};

const isApiUnavailableError = (err: unknown) => {
  if (!axios.isAxiosError(err)) {
    return false;
  }

  // 404 means route is unavailable in client-only runs.
  // No response usually means server is unreachable in dev.
  return err.response?.status === 404 || !err.response;
};

export const usePersonalization = () => {
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const savePersonalization = useCallback(
    async (data: Partial<PersonalizationData>) => {
      setLoading(true);
      setError(null);
      try {
        if (personalizationApiUnavailable) {
          const local = writeLocalPersonalization(data);
          setLoading(false);
          return local;
        }

        const response = await axios.post(
          withApiBase("/api/personalization/SavePersonalization"),
          data
        );

        // Keep local cache in sync for faster reloads.
        writeLocalPersonalization({ ...data, ...response.data });
        setLoading(false);
        return response.data;
      } catch (err) {
        if (isApiUnavailableError(err)) {
          personalizationApiUnavailable = true;
          const local = writeLocalPersonalization(data);
          setLoading(false);
          return local;
        }

        const errorMsg =
          err instanceof Error ? err.message : "Failed to save personalization";
        setError(errorMsg);
        setLoading(false);
        throw err;
      }
    },
    []
  );

  const getPersonalization = useCallback(
    async (userId: string, reportId: string) => {
      setLoading(true);
      setError(null);
      try {
        if (personalizationApiUnavailable) {
          const local = readLocalPersonalization(userId, reportId);
          setLoading(false);
          return local;
        }

        const response = await axios.get(
          `${withApiBase(
            "/api/personalization/GetPersonalization"
          )}?userId=${encodeURIComponent(
            userId
          )}&reportId=${encodeURIComponent(reportId)}`
        );

        if (response.data) {
          writeLocalPersonalization({
            userId,
            reportId,
            ...response.data,
          });
        }
        setLoading(false);
        return response.data;
      } catch (err) {
        if (isApiUnavailableError(err)) {
          personalizationApiUnavailable = true;
          const local = readLocalPersonalization(userId, reportId);
          setLoading(false);
          return local;
        }

        const errorMsg =
          err instanceof Error
            ? err.message
            : "Failed to get personalization";
        setError(errorMsg);
        setLoading(false);
        return null;
      }
    },
    []
  );

  const deletePersonalization = useCallback(async (id: string) => {
    setLoading(true);
    setError(null);
    try {
      if (!personalizationApiUnavailable) {
        await axios.delete(
          `${withApiBase(
            "/api/personalization/DeletePersonalization"
          )}?id=${encodeURIComponent(id)}`
        );
      }
      setLoading(false);
    } catch (err) {
      if (isApiUnavailableError(err)) {
        personalizationApiUnavailable = true;
        setLoading(false);
        return;
      }

      const errorMsg =
        err instanceof Error ? err.message : "Failed to delete personalization";
      setError(errorMsg);
      setLoading(false);
      throw err;
    }
  }, []);

  return {
    savePersonalization,
    getPersonalization,
    deletePersonalization,
    loading,
    error,
  };
};
