!/usr/bin/env bash
set -euo pipefail

# Git Bash on Windows otherwise rewrites container paths before Docker receives them.
export MSYS_NO_PATHCONV="${MSYS_NO_PATHCONV:-1}"

mode="${1:-smoke}"
if [[ "$mode" != "smoke" && "$mode" != "--fontless" && "$mode" != "--capacity" && "$mode" != "--metadata" ]]; then
  echo "usage: $0 [smoke|--fontless|--capacity|--metadata]" >&2
  exit 2
fi

image_name="${BING_OFFICES_VERIFY_IMAGE:-bing-offices-verify:local}"
fontless_image="${BING_OFFICES_FONTLESS_IMAGE:-bing-offices-verify:fontless}"
memory_limit="${BING_OFFICES_VERIFY_MEMORY:-4g}"
evidence_dir="${BING_OFFICES_DOCKER_EVIDENCE:-$PWD/artifacts/docker}"
mkdir -p "$evidence_dir"

docker build --file docker/verify/Dockerfile --tag "$image_name" .
docker version --format '{{json .}}' >"$evidence_dir/docker-version.json"
docker image inspect --format '{"id":{{json .Id}},"repoDigests":{{json .RepoDigests}},"os":{{json .Os}},"architecture":{{json .Architecture}}}' \
  "$image_name" >"$evidence_dir/image-identity.json"
docker run --rm "$image_name" >"$evidence_dir/dotnet-info.txt"

if [[ "$mode" == "--metadata" ]]; then
  exit 0
fi

common_args=(
  --rm
  --read-only
  --memory="$memory_limit"
  --user 10001:10001
  --volume "$PWD:/source:ro"
  --tmpfs /tmp:rw,exec,size=2g,uid=10001,gid=10001
  --tmpfs /home/bingoffices:rw,exec,size=1g,uid=10001,gid=10001
  --tmpfs /workspace:rw,exec,size=4g,uid=10001,gid=10001
  --workdir /workspace
  --entrypoint /bin/bash
)

if [[ "$mode" == "smoke" || "$mode" == "--fontless" ]]; then
  smoke_log="$evidence_dir/docker-smoke.log"
  : >"$smoke_log"
  docker build --file docker/verify/Dockerfile --target fontless --tag "$fontless_image" . \
    2>&1 | tee -a "$smoke_log"
  docker image inspect --format '{"id":{{json .Id}},"repoDigests":{{json .RepoDigests}},"os":{{json .Os}},"architecture":{{json .Architecture}}}' \
    "$fontless_image" >"$evidence_dir/fontless-image-identity.json"

  docker run "${common_args[@]}" "$image_name" -lc '
    set -euo pipefail
    dotnet --list-runtimes
    fc-list : family | grep -E "DejaVu|Noto Sans CJK" | head -n 2
  ' 2>&1 | tee -a "$smoke_log"

  docker run "${common_args[@]}" "$fontless_image" -lc '
    set -euo pipefail
    test ! -x /usr/bin/fc-list
    font_file_count="$(find /usr/share/fonts /usr/local/share/fonts -type f | wc -l)"
    test "$font_file_count" -eq 0
    echo "fontless-font-files=$font_file_count"
    tar -C /source \
      --exclude=".git" \
      --exclude=".vs" \
      --exclude=".workflow" \
      --exclude=".codex" \
      --exclude=".agents" \
      --exclude="artifacts" \
      --exclude="output" \
      --exclude="__pycache__" \
      --exclude=".tmp-*" \
      --exclude="*/bin" \
      --exclude="*/obj" \
      -cf - . | tar -C /workspace -xf -
    mkdir -p /home/bingoffices/.dotnet /home/bingoffices/.nuget/packages
    export DOTNET_CLI_HOME=/home/bingoffices/.dotnet
    export NUGET_PACKAGES=/home/bingoffices/.nuget/packages
    probe_project=build/StreamingProbe/Bing.Offices.StreamingProbe.csproj
    dotnet restore "$probe_project" --ignore-failed-sources -v:minimal
    dotnet build "$probe_project" --no-restore -c Release -f net8.0 -v:minimal -m:1
    probe_dll=build/StreamingProbe/bin/Release/net8.0/Bing.Offices.StreamingProbe.dll
    test -f "$probe_dll"
    probe_result="$(dotnet "$probe_dll" --npoi-font-probe)"
    printf "FONTLESS_NPOI_RESULT=%s\n" "$probe_result"
  ' 2>&1 | tee -a "$smoke_log"

  sed -n 's/^FONTLESS_NPOI_RESULT=//p' "$smoke_log" >"$evidence_dir/fontless-npoi-result.json"
  test "$(wc -l <"$evidence_dir/fontless-npoi-result.json")" -eq 1
  if [[ "$mode" == "--fontless" ]]; then
    exit 0
  fi

  docker run "${common_args[@]}" "$image_name" -lc '
    set -euo pipefail
    tar -C /source \
      --exclude=".git" \
      --exclude=".vs" \
      --exclude=".workflow" \
      --exclude=".codex" \
      --exclude=".agents" \
      --exclude="artifacts" \
      --exclude="output" \
      --exclude="__pycache__" \
      --exclude=".tmp-*" \
      --exclude="*/bin" \
      --exclude="*/obj" \
      -cf - . | tar -C /workspace -xf -
    mkdir -p /home/bingoffices/.dotnet /home/bingoffices/.nuget/packages
    export DOTNET_CLI_HOME=/home/bingoffices/.dotnet
    export NUGET_PACKAGES=/home/bingoffices/.nuget/packages
    provider_project=tests/Bing.Offices.ProviderContract.Tests/Bing.Offices.ProviderContract.Tests.csproj
    npoi_project=tests/Bing.Offices.Npoi.Tests/Bing.Offices.Npoi.Tests.csproj
    closed_xml_project=tests/Bing.Offices.ClosedXml.Tests/Bing.Offices.ClosedXml.Tests.csproj
    spread_project=tests/Bing.Offices.SpreadCheetah.Tests/Bing.Offices.SpreadCheetah.Tests.csproj
    for project in "$provider_project" "$npoi_project" "$closed_xml_project" "$spread_project"; do
      dotnet restore "$project" --ignore-failed-sources -v:minimal
    done
    for framework in net6.0 net8.0; do
      dotnet test "$provider_project" --no-restore -c Release --framework "$framework" \
        --filter FullyQualifiedName~SheetContentContractTest -v:minimal
      dotnet test "$npoi_project" --no-restore -c Release --framework "$framework" \
        --filter FullyQualifiedName~NpoiReportContractTest -v:minimal
      dotnet test "$closed_xml_project" --no-restore -c Release --framework "$framework" \
        --filter 'FullyQualifiedName~ClosedXmlReportContractTest|FullyQualifiedName~NamedListValidationTest' -v:minimal
      dotnet test "$spread_project" --no-restore -c Release --framework "$framework" -v:minimal
    done
  ' 2>&1 | tee -a "$smoke_log"
  exit 0
fi

capacity_log="$evidence_dir/streaming-capacity.log"
capacity_jsonl="$evidence_dir/streaming-capacity.jsonl"
: >"$capacity_log"

docker run "${common_args[@]}" "$image_name" -lc '
  set -euo pipefail
  tar -C /source \
    --exclude=".git" \
    --exclude=".vs" \
    --exclude=".workflow" \
    --exclude=".codex" \
    --exclude=".agents" \
    --exclude="artifacts" \
    --exclude="output" \
    --exclude="__pycache__" \
    --exclude=".tmp-*" \
    --exclude="*/bin" \
    --exclude="*/obj" \
    -cf - . | tar -C /workspace -xf -
  mkdir -p /home/bingoffices/.dotnet /home/bingoffices/.nuget/packages /workspace/capacity
  export DOTNET_CLI_HOME=/home/bingoffices/.dotnet
  export NUGET_PACKAGES=/home/bingoffices/.nuget/packages
  probe_project=build/StreamingProbe/Bing.Offices.StreamingProbe.csproj
  dotnet restore "$probe_project" --ignore-failed-sources -v:minimal
  dotnet build "$probe_project" --no-restore -c Release -f net8.0 -v:minimal -m:1
  probe_dll=build/StreamingProbe/bin/Release/net8.0/Bing.Offices.StreamingProbe.dll
  test -f "$probe_dll"
  for rows in 100000 500000 1000000; do
    for batch_size in 100 1000 10000; do
      output="/workspace/capacity/rows-${rows}-batch-${batch_size}.xlsx"
      probe_tmp="/tmp/streaming-${rows}-${batch_size}"
      mkdir -p "$probe_tmp"
      TMPDIR="$probe_tmp" dotnet "$probe_dll" \
        --rows "$rows" --batch-size "$batch_size" --output "$output"
      test -s "$output"
      rm -f "$output"
      find "$probe_tmp" -mindepth 1 -delete
      rmdir "$probe_tmp"
    done
  done
' 2>&1 | tee "$capacity_log"

grep -E '^\{"status":"success","rows":' "$capacity_log" >"$capacity_jsonl"
test "$(wc -l <"$capacity_jsonl")" -eq 9
