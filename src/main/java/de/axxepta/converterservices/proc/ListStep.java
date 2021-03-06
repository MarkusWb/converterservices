package de.axxepta.converterservices.proc;

import de.axxepta.converterservices.utils.IOUtils;
import de.axxepta.converterservices.utils.StringUtils;

import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

class ListStep extends Step {

    ListStep(String name, Object input, Object output, Object additional, boolean stopOnError, String... params) {
        super(name, input, output, additional, stopOnError, params);
    }

    @Override
    Pipeline.StepType getType() {
        return Pipeline.StepType.LIST;
    }

    @Override
    Object execAction(final List<String> inputFiles, final String... parameters)
            throws Exception
    {
        boolean xml = false;
        String charset = "UTF-8";
        for (String parameter : parameters) {
            String[] parts = parameter.split(" *= *");
            if (parts.length > 1) {
                switch (parts[0].toLowerCase()) {
                    case "xml":
                        if (parts[1].toLowerCase().equals("true")) {
                            xml = true;
                        }
                        break;
                    case "char": case "charset":
                        charset = parts[1];
                        break;
                }
            }
        }
        String outputFile = IOUtils.pathCombine(pipe.getWorkPath(),
                StringUtils.isNoStringOrEmpty(output) ? "list_step" + pipe.getCounter() + (xml ? ".xml" : ".txt") : (String) output);

        if (xml) {
            List<String> xmlLines = new ArrayList<>();
            xmlLines.add("<list>");
            xmlLines.addAll(inputFiles.stream().map(s -> "  <item>" + s + "</item>").collect(Collectors.toList()));
            xmlLines.add("</list>");
            IOUtils.saveStringArrayToFile(xmlLines, outputFile, true, charset);
        } else {
            IOUtils.saveStringArrayToFile(inputFiles, outputFile, true, charset);
        }

        pipe.addGeneratedFile(outputFile);
        return singleFileList(outputFile);
    }

    @Override
    protected boolean assertParameter(final Parameter paramType, final Object param) {
        return !paramType.equals(Parameter.INPUT) || assertStandardInput(param);
    }
}
