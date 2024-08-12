import { checkScript, preAndAllocate, allocateCSVFanAndHib } from "./functions/master";
import { browseXlsFile, browseCSVFolder, browseDesktopFolder } from "./functions/setup";
import { getProposedThresholds, updateProposedThresholds, updateProposedThreshodsGreen, updateProposedThresholds10 } from "./functions/allocationThresholds";
import { paths } from "./utils/utils";
const inputPath = paths.inputPath;

(async () => {
    // await checkScript();
    // await preAndAllocate();
    // await allocateCSVFanAndHib();
    
    // await browseXlsFile();
    // await browseCSVFolder();
    // await browseDesktopFolder();

    // await getProposedThresholds(inputPath, true); // not working
    // await updateProposedThresholds(inputPath);
    // await updateProposedThreshodsGreen(inputPath);
    // await updateProposedThresholds10(inputPath);
})();