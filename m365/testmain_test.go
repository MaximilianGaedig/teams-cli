//go:build !integration

package m365

import (
	"os"
	"testing"

	"github.com/MaximilianGaedig/teams/internal/testenv"
)

func TestMain(m *testing.M) { os.Exit(testenv.Run(m)) }
