package com.example.aiexcel.config;


/**
 * Centralized environment helper. 统一从此类读取 API Key / Base URL / Model 等
 */
public final class EnvFile {

    private EnvFile() {}

    // Runtime overrides set via API (e.g., /api/sendapi). Kept in-memory; not persisted.
    private static volatile String runtimeApiKey = null;
    private static volatile String runtimeBaseUrl = null;

    public static void setRuntimeApiKey(String key) {
        runtimeApiKey = isPresent(key) ? key.trim() : null;
        if (runtimeApiKey != null) {
            System.setProperty("qwen.api.api-key", runtimeApiKey);
        }
    }

    public static void setRuntimeBaseUrl(String baseUrl) {
        runtimeBaseUrl = isPresent(baseUrl) ? baseUrl.trim() : null;
        if (runtimeBaseUrl != null) {
            System.setProperty("qwen.api.base-url", runtimeBaseUrl);
        }
    }

    public static String getApiKey() {
        if (isPresent(runtimeApiKey)) return runtimeApiKey.trim();
        String key = System.getenv("DASHSCOPE_API_KEY");
        if (isPresent(key)) return key.trim();

        key = System.getenv("QWEN_API_KEY");
        if (isPresent(key)) return key.trim();

        key = System.getProperty("qwen.api.api-key");
        if (isPresent(key)) return key.trim();

        // fall back to uppercase property if set by some loaders
        key = System.getProperty("QWEN_API_KEY");
        if (isPresent(key)) return key.trim();

        return null;
    }

    public static String getBaseUrl() {
        if (isPresent(runtimeBaseUrl)) return runtimeBaseUrl.trim();
        String b = System.getProperty("qwen.api.base-url");
        if (isPresent(b)) return b.trim();

        b = System.getenv("QWEN_API_BASE_URL");
        if (isPresent(b)) return b.trim();

        return "https://dashscope.aliyuncs.com/compatible-mode/v1";
    }

    public static String getDefaultModel() {
        String m = System.getProperty("qwen.api.default-model");
        if (isPresent(m)) return sanitizeModel(m);

        m = System.getenv("QWEN_MODEL_NAME");
        if (isPresent(m)) return sanitizeModel(m);

        return "qwen-max";
    }

    public static boolean isDev() {
        String profile = System.getProperty("spring.profiles.active");
        if (isPresent(profile) && "dev".equalsIgnoreCase(profile.trim())) return true;
        String env = System.getenv("ENV");
        if (isPresent(env) && "dev".equalsIgnoreCase(env.trim())) return true;
        String localDev = System.getenv("LOCAL_DEV");
        return isPresent(localDev) && "true".equalsIgnoreCase(localDev.trim());
    }

    public static String mask(String key) {
        if (key == null) return "<none>";
        String t = key.trim();
        if (t.length() <= 10) return "***";
        return t.substring(0, Math.min(6, t.length())) + "***" + t.substring(Math.max(t.length() - 4, 0));
    }

    private static boolean isPresent(String s) {
        return s != null && !s.trim().isEmpty();
    }

    private static String sanitizeModel(String m) {
        if (m == null) return null;
        String out = m.trim();
        int hashIdx = out.indexOf('#');
        if (hashIdx >= 0) out = out.substring(0, hashIdx).trim();
        int slashIdx = out.indexOf("//");
        if (slashIdx >= 0) out = out.substring(0, slashIdx).trim();
        return out;
    }
}
