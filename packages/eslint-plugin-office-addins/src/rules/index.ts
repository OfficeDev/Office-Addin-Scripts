import loadObjectBeforeRead from "./load-object-before-read";
import noOfficeInitialize from "./no-office-initialize";
import noEmptyLoad from "./no-empty-load";
import noContextSyncInLoop from "./no-context-sync-in-loop";
import testForNullUsingIsNullObject from "./test-for-null-using-isNullObject";

export default {
  "load-object-before-read": loadObjectBeforeRead,
  "no-context-sync-in-loop": noContextSyncInLoop,
  "no-empty-load": noEmptyLoad,
  "no-office-initialize": noOfficeInitialize,
  "test-for-null-using-isNullObject": testForNullUsingIsNullObject,
};
