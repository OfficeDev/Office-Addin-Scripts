import callSyncAfterLoad from "./call-sync-after-load";
import callSyncBeforeRead from "./call-sync-before-read";
import loadObjectBeforeRead from "./load-object-before-read";
import noContextSyncInLoop from "./no-context-sync-in-loop";
import noEmptyLoad from "./no-empty-load";
import noNavigationalLoad from "./no-navigational-load";
import noOfficeInitialize from "./no-office-initialize";
import testForNullUsingIsNullObject from "./test-for-null-using-isNullObject";

export default {
  "call-sync-before-read": callSyncBeforeRead,
  "load-object-before-read": loadObjectBeforeRead,
  "call-sync-after-load": callSyncAfterLoad,
  "no-context-sync-in-loop": noContextSyncInLoop,
  "no-empty-load": noEmptyLoad,
  "no-navigational-load": noNavigationalLoad,
  "no-office-initialize": noOfficeInitialize,
  "test-for-null-using-isNullObject": testForNullUsingIsNullObject,
};
