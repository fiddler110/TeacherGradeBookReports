#!/bin/sh
set -e
# Fix ownership of the persistent /data volume so appuser can write to it.
# This handles the case where the volume was previously created by a
# root-running container.  The chown is a no-op when permissions are
# already correct, so there is no performance penalty on subsequent starts.
chown -R appuser:appgroup /data
exec gosu appuser "$@"
