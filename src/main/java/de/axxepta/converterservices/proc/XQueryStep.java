package de.axxepta.converterservices.proc;

import de.axxepta.converterservices.tools.Saxon;
import de.axxepta.converterservices.utils.IOUtils;
import de.axxepta.converterservices.utils.StringUtils;
import org.w3c.dom.Document;

import java.util.List;

class XQueryStep extends Step {

    XQueryStep(Object input, Object output, Object additional, String... params) {
        super(input, output, additional, params);
    }

    Pipeline.StepType getType() {
        return Pipeline.StepType.XQUERY;
    }

    @Override
    Object execAction(final Pipeline pipe, final List<String> inputFiles, final Object additionalInput, final String... parameters) throws Exception {
        String queryFile = pipedPath(additionalInput, pipe);
        String query = IOUtils.readTextFile(queryFile);
        Object queryOutput = pipe.saxonXQuery(query,
                input.equals(Saxon.XQUERY_NO_CONTEXT) ? (String) input : inputFiles.get(0),
                parameters);
        if (queryOutput instanceof Document) {
            String outputFile = StringUtils.isEmpty(output) ?
                    IOUtils.pathCombine(pipe.getWorkPath(),"output_step" + pipe.getCounter() + ".xml") :
                    IOUtils.pathCombine(pipe.getWorkPath(), (String) output);
            Saxon.saveDOM((Document) queryOutput, outputFile, "ISO-8859-1");
            pipe.addGeneratedFile(outputFile);
            actualOutput = outputFile;
            return singleFileList(outputFile);
        } else {
            // ToDo: check cases, assure correct feeding in pipe
            actualOutput = queryOutput;
            return queryOutput;
        }
    }

    @Override
    protected boolean assertParameter(final Parameter paramType, final Object param) {
        switch (paramType) {
            case INPUT:
                break;
            case OUTPUT:
                break;
            case ADDITIONAL:
                break;
            case PARAMS:
                break;
            default: return true;
        }
        return true;
        //if (StringUtils.isEmpty(param) && paramType.equals(Parameter.ADDITIONAL))
        //    return false;
        //return (param instanceof String) || ((param instanceof List) && ((List) param).get(0) instanceof String);
    }
}
