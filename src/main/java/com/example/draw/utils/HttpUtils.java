package com.example.draw.utils;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;

import javax.net.ssl.HttpsURLConnection;
import javax.net.ssl.SSLContext;
import javax.net.ssl.TrustManager;
import javax.net.ssl.X509TrustManager;
import java.io.*;
import java.net.HttpURLConnection;
import java.net.MalformedURLException;
import java.net.URL;
import java.security.cert.X509Certificate;
import java.util.Map;
import java.util.UUID;

@Component
public class HttpUtils {

    private static Logger logger = LoggerFactory.getLogger(HttpUtils.class);

    public static String doGet(String requestURL, Map<String, String> params, String token) {
        if (params != null) {
            requestURL = requestURL + "?";
            for (Map.Entry<String, String> entry : params.entrySet()) {
                String value = entry.getValue();
                if (value == null) {
                    value = "";
                }
                requestURL = requestURL + entry.getKey() + "=" + entry.getValue() + "&";
            }
            requestURL = requestURL.substring(0, requestURL.length() - 1);
        }
        logger.info("HttpUtils doGet request url ======>> : " + requestURL);
        HttpURLConnection connection = null;
        InputStream is = null;
        BufferedReader br = null;
        String result = null;// 返回结果字符串
        try {
            trustAllHosts();
            // 创建远程url连接对象
            URL url = new URL(requestURL);
            // 通过请求地址判断请求类型(http或者是https)
            if ("https".equals(url.getProtocol().toLowerCase())) {
                HttpsURLConnection https = (HttpsURLConnection) url.openConnection();
                // 通过远程url连接对象打开连接
//                https.setHostnameVerifier(new HostnameVerifier() {
//                    @Override
//                    public boolean verify(String arg0, SSLSession arg1) {
//                        return true;
//                    }
//                });
                connection = https;
            } else {
                connection = (HttpURLConnection) url.openConnection();
            }
            // 设置连接方式：get
            connection.setRequestMethod("GET");
            // 设置连接主机服务器的超时时间：10秒
            connection.setConnectTimeout(30000);
            // 设置读取远程返回的数据时间：10秒
            connection.setReadTimeout(30000);
            // 默认值为：false，当向远程服务器传送数据/写数据时，需要设置为true
            connection.setDoOutput(true);
            // 默认值为：true，当前向远程服务读取数据
            connection.setDoInput(true);
            // 设置字符集
            connection.setRequestProperty("Charset", "UTF-8");
            // 设置传入参数的格式
            connection.setRequestProperty("Content-Type", "application/json");
            // 设置鉴权信息
            if (token != null) {
                connection.setRequestProperty("Authorization", token);
            }
            // 发送请求
            connection.connect();
            logger.info("HttpUtils doGet requestUrl=" + requestURL + ", responseCode=" + connection.getResponseCode());
            // 通过connection连接，获取输入流
            if (connection.getResponseCode() == HttpURLConnection.HTTP_OK) {
                is = connection.getInputStream();
                // 封装输入流is，并指定字符集
                br = new BufferedReader(new InputStreamReader(is, "UTF-8"));
                // 存放数据
                StringBuffer sbf = new StringBuffer();
                String temp = null;
                while ((temp = br.readLine()) != null) {
                    sbf.append(temp);
                    sbf.append("\r\n");
                }
                result = sbf.toString();
            } else {
                is = connection.getErrorStream();
                // 对输入流对象进行包装，指定charset
                br = new BufferedReader(new InputStreamReader(is, "UTF-8"));
                StringBuilder builder = new StringBuilder();
                String temp = null;
                // 循环遍历一行一行读取数据
                while ((temp = br.readLine()) != null) {
                    builder.append(temp);
                }
                logger.error("HttpUtils doGet exception:{}", builder.toString());
                throw new RuntimeException(connection.getResponseCode() + "-" + builder.toString());
            }
        } catch (MalformedURLException e) {
            e.printStackTrace();
            logger.error("HttpUtils doGet exception", e);
        } catch (IOException e) {
            e.printStackTrace();
            logger.error("HttpUtils doGet exception", e);
        } finally {
            // 关闭资源
            if (null != br) {
                try {
                    br.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }

            if (null != is) {
                try {
                    is.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            // 关闭远程连接
            connection.disconnect();
        }

        return result;
    }

    public static InputStream doGet2(String requestURL, Map<String, String> params, String token) {
        if (params != null) {
            requestURL = requestURL + "?";
            for (Map.Entry<String, String> entry : params.entrySet()) {
                String value = entry.getValue();
                if (value == null) {
                    value = "";
                }
                requestURL = requestURL + entry.getKey() + "=" + entry.getValue() + "&";
            }
            requestURL = requestURL.substring(0, requestURL.length() - 1);
        }
        logger.info("HttpUtils doGet request url ======>> : " + requestURL);
        HttpURLConnection connection = null;
        InputStream is = null;
        BufferedReader br = null;
        try {
            trustAllHosts();
            // 创建远程url连接对象
            URL url = new URL(requestURL);
            // 通过请求地址判断请求类型(http或者是https)
            if ("https".equals(url.getProtocol().toLowerCase())) {
                HttpsURLConnection https = (HttpsURLConnection) url.openConnection();
                connection = https;
            } else {
                connection = (HttpURLConnection) url.openConnection();
            }
            // 设置连接方式：get
            connection.setRequestMethod("GET");
            // 设置连接主机服务器的超时时间：10秒
            connection.setConnectTimeout(30000);
            // 设置读取远程返回的数据时间：10秒
            connection.setReadTimeout(30000);
            // 默认值为：false，当向远程服务器传送数据/写数据时，需要设置为true
            connection.setDoOutput(true);
            // 默认值为：true，当前向远程服务读取数据
            connection.setDoInput(true);
            // 设置字符集
//            connection.setRequestProperty("Charset", "UTF-8");
            // 设置传入参数的格式
//            connection.setRequestProperty("Content-Type", "application/json");
            // 设置鉴权信息
            if (token != null) {
                connection.setRequestProperty("Authorization", token);
            }
            // 发送请求
            connection.connect();
            logger.info("HttpUtils doGet requestUrl=" + requestURL + ", responseCode=" + connection.getResponseCode());
            // 通过connection连接，获取输入流
            if (connection.getResponseCode() == HttpURLConnection.HTTP_OK) {
                is = connection.getInputStream();
                return is;
            } else {
                is = connection.getErrorStream();
                // 对输入流对象进行包装，指定charset
                br = new BufferedReader(new InputStreamReader(is, "UTF-8"));
                StringBuilder builder = new StringBuilder();
                String temp = null;
                // 循环遍历一行一行读取数据
                while ((temp = br.readLine()) != null) {
                    builder.append(temp);
                }
                logger.error("HttpUtils doGet exception:{}", builder.toString());
                throw new RuntimeException(connection.getResponseCode() + "-" + builder.toString());
            }
        } catch (MalformedURLException e) {
            e.printStackTrace();
            logger.error("HttpUtils doGet exception", e);
        } catch (IOException e) {
            e.printStackTrace();
            logger.error("HttpUtils doGet exception", e);
        } finally {
            // 关闭资源
            if (null != br) {
                try {
                    br.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }

//            if (null != is) {
//                try {
//                    is.close();
//                } catch (IOException e) {
//                    e.printStackTrace();
//                }
//            }
            // todo 111 待优化
            // 关闭远程连接
//            connection.disconnect();
        }
        return null;

//        return result;
    }


    public static String doPost(String requestURL, String params, String token) {
        HttpURLConnection connection = null;
        InputStream is = null;
        OutputStream os = null;
        BufferedReader br = null;
        String result = null;
        try {
            trustAllHosts();
            // 创建远程url连接对象
            URL url = new URL(requestURL);
            // 通过请求地址判断请求类型(http或者是https)
            if ("https".equals(url.getProtocol().toLowerCase())) {
                HttpsURLConnection https = (HttpsURLConnection) url.openConnection();
                // 通过远程url连接对象打开连接
//                https.setHostnameVerifier(new HostnameVerifier() {
//                    @Override
//                    public boolean verify(String arg0, SSLSession arg1) {
//                        return true;
//                    }
//                });
                connection = https;
            } else {
                connection = (HttpURLConnection) url.openConnection();
            }
            // 设置连接请求方式
            connection.setRequestMethod("POST");
            // 设置连接主机服务器超时时间：10s
            connection.setConnectTimeout(30000);
            // 设置读取主机服务器返回数据超时时间：10s
            connection.setReadTimeout(30000);
            // 默认值为：false，当向远程服务器传送数据/写数据时，需要设置为true
            connection.setDoOutput(true);
            // 默认值为：true，当前向远程服务读取数据
            connection.setDoInput(true);
            // 设置字符集
            connection.setRequestProperty("Charset", "UTF-8");
            // 设置传入参数的格式
            connection.setRequestProperty("Content-Type", "application/json");
            byte[] bytes = params.getBytes("UTF-8");
            connection.setRequestProperty("Content-length", String.valueOf(bytes.length));
            // 设置鉴权信息
            if (token != null) {
                connection.setRequestProperty("Authorization", token);
            }
            // 通过连接对象获取一个输出流
            os = connection.getOutputStream();
            // 通过输出流对象将参数写出去/传输出去，它是通过字节数组写出的
            os.write(bytes);
            // 通过连接对象获取一个输入流，向远程读取
            logger.info("HttpUtils doPost requestUrl=" + requestURL + ", responseCode=" + connection.getResponseCode() + " ,requestParameter=" + params);
            if (connection.getResponseCode() == HttpURLConnection.HTTP_OK) {
                is = connection.getInputStream();
                // 对输入流对象进行包装，指定charset
                br = new BufferedReader(new InputStreamReader(is, "UTF-8"));
                StringBuilder builder = new StringBuilder();
                String temp = null;
                // 循环遍历一行一行读取数据
                while ((temp = br.readLine()) != null) {
                    builder.append(temp);
                }
                result = builder.toString();
            } else {
                is = connection.getErrorStream();
                // 对输入流对象进行包装，指定charset
                br = new BufferedReader(new InputStreamReader(is, "UTF-8"));
                StringBuilder builder = new StringBuilder();
                String temp = null;
                // 循环遍历一行一行读取数据
                while ((temp = br.readLine()) != null) {
                    builder.append(temp);
                }
                logger.error("HttpUtils doPost exception:{}", builder.toString());
                throw new RuntimeException(connection.getResponseCode() + "-" + builder.toString());
            }
        } catch (MalformedURLException e) {
            e.printStackTrace();
            logger.error("HttpUtils doPost exception", e);
        } catch (IOException e) {
            e.printStackTrace();
            logger.error("HttpUtils doPost exception", e);
        } finally {
            // 关闭资源
            if (br != null) {
                try {
                    br.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (os != null) {
                try {
                    os.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (is != null) {
                try {
                    is.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            // 断开与远程地址url的连接
            connection.disconnect();
        }
        return result;
    }

    static String doPost2(String url, Map<String, String> params, Map<String, byte[]> fileMap, String token) throws IOException {
        HttpURLConnection conn = null;
        DataOutputStream outputStream = null;
        BufferedReader reader = null;
        try {
            StringBuilder sb = new StringBuilder();
            String BOUNDARY = UUID.randomUUID().toString();
            String PREFIX = "--", LINEND = "\r\n";
            String MULTIPART_FROM_DATA = "multipart/form-data";
            String CHARSET = "UTF-8";

            URL uri = new URL(url);
            conn = (HttpURLConnection) uri.openConnection();
            conn.setReadTimeout(600000);
            conn.setDoInput(true);
            conn.setDoOutput(true);
            conn.setUseCaches(false);
            conn.setRequestMethod("POST");
            conn.setRequestProperty("connection", "keep-alive");
            conn.setRequestProperty("Charset", "UTF-8");
            conn.setRequestProperty("Content-Type", MULTIPART_FROM_DATA + ";boundary=" + BOUNDARY);
            if (token != null) {
                conn.setRequestProperty("Authorization", token);
            }

            // 文本类型参数
            StringBuilder sbParams = new StringBuilder();
            for (Map.Entry<String, String> entry : params.entrySet()) {
                sbParams.append(PREFIX);
                sbParams.append(BOUNDARY);
                sbParams.append(LINEND);
                sbParams.append("Content-Disposition: form-data; name=\""
                        + entry.getKey() + "\"" + LINEND);
                sbParams.append("Content-Type: text/plain; charset=" + CHARSET + LINEND);
                sbParams.append(LINEND);
                sbParams.append(entry.getValue());
                sbParams.append(LINEND);
            }

            outputStream = new DataOutputStream(conn.getOutputStream());
            outputStream.write(sbParams.toString().getBytes());
            InputStream in = null;
            // 文件类型数据
            if (fileMap != null) {
                for (Map.Entry<String, byte[]> file : fileMap.entrySet()) {
                    StringBuilder sbFile = new StringBuilder();
                    sbFile.append(PREFIX);
                    sbFile.append(BOUNDARY);
                    sbFile.append(LINEND);
                    sbFile.append("Content-Disposition: form-data; name=\"file\"; filename=\""
                            + file.getKey() + "\"" + LINEND);
                    sbFile.append("Content-Type:" + "application/octet-stream;UTF-8"
                            + LINEND);
                    sbFile.append(LINEND);
                    outputStream.write(sbFile.toString().getBytes());
                    outputStream.write(file.getValue());
                    outputStream.write(LINEND.getBytes());
                }
            }

            byte[] endData = (PREFIX + BOUNDARY + PREFIX + LINEND).getBytes();
            outputStream.write(endData);
            outputStream.flush();

            int code = conn.getResponseCode();
//            if (code == 200) {
            reader = new BufferedReader(new InputStreamReader(conn.getInputStream(), "UTF-8"));
            String line = null;
            while ((line = reader.readLine()) != null) {
                sb.append(line).append("\n");
            }
//            }
            return sb.toString();
        } catch (IOException e) {
            logger.error("=====AliYunOssUtil.doPost====方法报错:{}", e);
            throw e;
        } finally {
            if (reader != null) {
                try {
                    reader.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (outputStream != null) {
                try {
                    outputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (conn != null) {
                conn.disconnect();
            }
        }
    }

    private static byte[] toByteArray(InputStream is) throws IOException {
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        byte[] buffer = new byte[1024];
        int n = 0;
        while (-1 != (n = is.read(buffer))) {
            os.write(buffer, 0, n);
        }
        return os.toByteArray();
    }

    private static void trustAllHosts() {
        // Create a trust manager that does not validate certificate chains
        TrustManager[] trustAllCerts = new TrustManager[]{
                new X509TrustManager() {
                    public X509Certificate[] getAcceptedIssuers() {
                        return new X509Certificate[]{};
                    }

                    public void checkClientTrusted(X509Certificate[] chain, String authType) {
                    }

                    public void checkServerTrusted(X509Certificate[] chain, String authType) {
                    }
                }
        };
        // Install the all-trusting trust manager
        try {
            SSLContext sc = SSLContext.getInstance("SSL");
            sc.init(null, trustAllCerts, new java.security.SecureRandom());
            HttpsURLConnection.setDefaultSSLSocketFactory(sc.getSocketFactory());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


}
