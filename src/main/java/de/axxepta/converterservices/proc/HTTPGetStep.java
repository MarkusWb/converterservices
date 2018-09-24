package de.axxepta.converterservices.proc;

import de.axxepta.converterservices.utils.HTTPUtils;
import de.axxepta.converterservices.utils.IOUtils;
import de.axxepta.converterservices.utils.StringUtils;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

class HTTPGetStep extends Step {

    HTTPGetStep(String name, Object input, Object output, Object additional, String... params) {
        super(name, input, output, additional, params);
    }

    @Override
    Pipeline.StepType getType() {
        return Pipeline.StepType.HTTP_GET;
    }

    @Override
    Object execAction(final Pipeline pipe, final List<String> inputFiles, final String... parameters) throws Exception {

        String server = "";
        String user = "";
        String pwd = "";
        int port = 80;
        boolean secure = false;
        String contentTypeString = "";

        for (String parameter : parameters) {
            String[] parts = parameter.split(" *= *");
            if (parts.length > 1) {
                switch (parts[0].toLowerCase()) {
                    case "server": case "host":
                        server = parts[1];
                        break;
                    case "user":
                        user = parts[1];
                        break;
                    case "pwd": case "password":
                        pwd = parts[1];
                        break;
                    case "secure": case "ssl":
                        if (parts[1].toLowerCase().equals("true")) {
                            secure = true;
                            if (port == 80) {
                                port = 443;
                            }
                        }
                        break;
                    case "port":
                        if (StringUtils.isInt(parts[1])) {
                            port = Integer.valueOf(parts[1]);
                        }
                        break;
                    case "content": case "contenttype": case "accept":
                        contentTypeString = parts[1];
                        break;
                }
            }
        }

        List<String> providedOutputNames = listifyOutput(pipe);
        List<String> downloadedFiles = new ArrayList<>();
        String inputPath = IOUtils.relativePath(inputFiles.get(0), pipe.getInputPath());

        try {
            // ToDo: handle multipart responses, eventually multiple get requests/inputs
            String outputFile = IOUtils.pathCombine(pipe.getWorkPath(), providedOutputNames.size() == 0 ?
                    "step_" + pipe.getCounter() + ".xml" : providedOutputNames.get(0));
            System.out.println("HTTP_GET output file " + outputFile);
            System.out.println("");
            List<String> responseFiles;
            if (contentTypeString.equals("")) {
                responseFiles = HTTPUtils.get(secure ? "https" : "http", server, port, inputPath, user, pwd, outputFile);
            } else {
                responseFiles = HTTPUtils.get(secure ? "https" : "http", server, port, inputPath, user, pwd, outputFile, contentTypeString);
            }
            downloadedFiles.addAll(responseFiles);
        } catch (IOException ex) {
            pipe.log(String.format("Error during HTTP GET to %s: %s",
                    (secure ? "https" : "http" + server + port + inputPath), ex.getMessage()));
            throw ex;
        }

        return downloadedFiles;
    }

    @Override
    boolean assertParameter(Parameter paramType, Object param) {
        return true;
    }
}