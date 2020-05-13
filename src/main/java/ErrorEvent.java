import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import lombok.experimental.Accessors;

@Getter
@Setter
@NoArgsConstructor
@Accessors(chain = true)
public class ErrorEvent {
    private String code;
    private EVENT_MODE mode;

    public ErrorEvent(String code, EVENT_MODE mode) {
        this.code = code;
        this.mode = mode;
    }
}
