import callSyncBeforeRead from "./call-sync-before-read";
import loadObjectBeforeRead from "./load-object-before-read";
import noContextSyncInLoop from "./no-context-sync-in-loop";
import noOfficeInitialize from "./no-office-initialize";
import testForNullUsingIsNullObject from "./test-for-null-using-isNullObject";

export default {
  "call-sync-before-read": callSyncBeforeRead,
  "load-object-before-read": loadObjectBeforeRead,
  "no-context-sync-in-loop": noContextSyncInLoop,
  "no-office-initialize": noOfficeInitialize,
  "test-for-null-using-isNullObject": testForNullUsingIsNullObject,
};
